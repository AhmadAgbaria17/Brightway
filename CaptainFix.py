import os
import json

from langchain_core.prompts import ChatPromptTemplate
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from langchain_openai import ChatOpenAI

def set_up():
    load_dotenv()
    driver = webdriver.Chrome()
    # html_link = input("enter you link please:\n")
    # driver.get(html_link)
    driver.get("file:///C:/Users/Ahmed/Downloads/ActionChainsEx.Html")
    llm = ChatOpenAI(
        model='gpt-3.5-turbo',
        temperature=0.1,
        max_tokens=500,
        api_key= os.getenv("OPENAI_API_KEY")
    )
    return driver , llm


def analyze_html_with_llm(html_content: str, llm):
    template = (
        "אני נותן לך דף HTML של דף אינטרנט {html}\n"
        "זהה את האלמנטים הפעילים בדף (קישוריים, כפתורים, טפסים) ותאר את הפעולות שניתן לבצע עליהם.\n"
        "החזר אך ורק JSON תקין בלבד (בלי טקסט נוסף). תמיין לפי סוג האלמנט (links, buttons, forms, inputs וכו')."
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

    raw_text = json.loads(raw_text)
    return raw_text





def main():
    driver,llm = set_up()
    try:
        html_content = driver.page_source
        operations = analyze_html_with_llm(html_content, llm)



    finally:
        driver.quit()

if __name__ == "__main__":
    main()