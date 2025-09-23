import os
import json
from textwrap import indent

from langchain_core.prompts import ChatPromptTemplate
from dotenv import load_dotenv

import pandas as pd
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from langchain_openai import ChatOpenAI

def set_up():
    load_dotenv()
    driver = webdriver.Chrome()
    # html_link = input("enter you link please:\n")
    # driver.get(html_link)
    driver.get("file:///C:/Users/Ahmed/Downloads/User%20Managment(1).html")
    llm = ChatOpenAI(
        model='gpt-3.5-turbo',
        temperature=0.1,
        max_tokens=3000,
        api_key= os.getenv("OPENAI_API_KEY")
    )
    return driver , llm


def analyze_html_with_llm(html_content: str, llm):
    template = (
        "אני נותן לך דף HTML של דף אינטרנט {html}\n"
        "זהה את כל האלמנטים הפעילים בדף (קישוריים, כפתורים, טפסים, אם יש עוד) ותאר את הפעולות שניתן לבצע עליהם.\n"
        "החזר אך ורק JSON תקין בלבד (בלי טקסט נוסף). תמיין לפי סוג האלמנט (links, buttons, forms, inputs,other וכו')."
    )

    prompt = ChatPromptTemplate.from_template(template)
    chain = prompt | llm

    try:
        response = chain.invoke({"html": html_content})
    except Exception as e:
        raise RuntimeError(f"Failed to invoke chain: {e}")

    if hasattr(response, "content"):
        raw_text = response.content
    else:
        raw_text = str(response)
    return raw_text


# generate a test cases
def generate_test_cases_by_llm(html_content:str , elements:str , llm):
    template =  """
        HTML: {html}
        ELEMENTS: {elements}   # e.g. ["/", "/login", "/contact", "/products"]
        Constraints:
        - Generate Suites: (Smoke, Navigation, Forms).
        - Produce 5-12 test cases total 
        - Each test case must include: id, title, priority (P0,P1,P2), steps (action/selector/value), expected(very important), tags (optional).
        - Allowed actions: goto, click, fill, assert_text, assert_element, assert_url, assert_status, wait, select, upload, hover, back, forward.
        - Use selectors that are likely to exist (ids, data-test attributes). If no selector possible, use the page-level step action (e.g., "assert_text" with selector "body").
        - Every step must be a small atomic action (one action per step).
        - Use consistent case for ids: <SUITE>-C<number> (e.g., SMOKE-C1).
        - Output must validate as JSON. Do not wrap the JSON in markdown or code fences.
        - If a page from the provided list doesn't exist, include a test case to assert 200 status for that page.
        - Please generate the full JSON test plan with all suites and cases in one reply, don’t truncate.
        Return the TestPlan JSON now.
    """
    prompt = ChatPromptTemplate.from_template(template)
    chain = prompt | llm


    try:
        response = chain.invoke({"html":html_content , "elements":elements})
    except Exception as e:
        raise RuntimeError(f"Failed to invoke chain: {e}")
    if hasattr(response, "content"):
        test_cases = response.content
    else:
        test_cases = str(response)

    test_cases = json.loads(test_cases)

    return test_cases


# export json file
def export_testplan_json_file(filename: str , content: dict):
    with open(filename, "w" , encoding="utf-8") as f:
        json.dump(content , f , indent=4 , ensure_ascii=False)

def export_testplan_to_excel( filename: str , content: dict):
    """
    Export testplan JSON -> Excel workbook with two sheets:
      - TestCases: one row per test case (high-level)
      - Steps: one row per step (detailed)
    """
    # Flatten test cases
    case_rows = []
    step_rows = []
    for suite in content.get("suites", []):
        suite_name = suite.get("name", "")
        for case in suite.get("cases", []):
            case_id = case.get("id", "")
            title = case.get("title", "")
            priority = case.get("priority", "")
            expected = case.get("expected", "")
            tags = ", ".join(case.get("tags", []))
            # Build steps text (multi-line)
            steps = case.get("steps", [])
            step_texts = []
            for i, s in enumerate(steps, start=1):
                # Represent step succinctly
                action = s.get("action", "")
                selector = s.get("selector") or s.get("value") or ""
                value = s.get("value") if "value" in s and s.get("selector") else ""
                # prefer explicit fields
                if s.get("selector") and "value" in s:
                    step_text = f"{i}. {action} selector={s.get('selector')} value={s.get('value')}"
                elif s.get("selector"):
                    step_text = f"{i}. {action} selector={s.get('selector')}"
                elif "value" in s:
                    step_text = f"{i}. {action} value={s.get('value')}"
                else:
                    step_text = f"{i}. {action}"
                step_texts.append(step_text)

                # add to detailed steps sheet
                step_rows.append({
                    "Suite": suite_name,
                    "Case ID": case_id,
                    "Step No": i,
                    "Action": action,
                    "Selector": s.get("selector", ""),
                    "Value": s.get("value", ""),
                    "Raw": json.dumps(s, ensure_ascii=False)
                })

            steps_cell = "\n".join(step_texts)
            case_rows.append({
                "Suite": suite_name,
                "Case ID": case_id,
                "Title": title,
                "Priority": priority,
                "Expected": expected,
                "Tags": tags,
                "Steps": steps_cell
            })

    # Create DataFrames
    df_cases = pd.DataFrame(case_rows)
    df_steps = pd.DataFrame(step_rows)

    # Write to Excel with two sheets
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_cases.to_excel(writer, sheet_name="TestCases", index=False)
        df_steps.to_excel(writer, sheet_name="Steps", index=False)


    # Auto-adjust column widths (openpyxl)
    try:
        from openpyxl import load_workbook
        wb = load_workbook(filename)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # determine max width per column
            dims = {}
            for row in ws.rows:
                for cell in row:
                    if cell.value is not None:
                        cell_value = str(cell.value)
                        dims[cell.column_letter] = max(dims.get(cell.column_letter, 0), len(cell_value))
            for col, value in dims.items():
                adjusted = min(value + 2, 60)  # cap width
                ws.column_dimensions[col].width = adjusted
        wb.save(filename)
    except Exception as e:
        # Non-fatal: still valid Excel, just not auto-sized
        print("Warning: could not auto-size columns:", e)





def main():
    driver,llm = set_up()
    try:
        html_content = driver.page_source
        elements = analyze_html_with_llm(html_content, llm)
        test_cases = generate_test_cases_by_llm(html_content,elements , llm)
        export_testplan_json_file("PlanJson.txt", test_cases)
        export_testplan_to_excel("planExcel.xlsx",test_cases )

    finally:
        driver.quit()

if __name__ == "__main__":
    main()