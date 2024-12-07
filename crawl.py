import ddddocr
import requests
import re
import json
from bs4 import BeautifulSoup
from tqdm import tqdm

JSESSIONID = "hcydsg_V4CAIBe2TOZHoz"
BASE_URL_1 = "https://zhjwxk.cic.tsinghua.edu.cn/xkBks.xkBksZytjb.do"
BASE_URL_2 = "https://zhjwxk.cic.tsinghua.edu.cn/xkBks.vxkBksJxjhBs.do"
CAPTCHA_URL = "https://zhjwxk.cic.tsinghua.edu.cn/login-jcaptcah.jpg?captchaflag=login1"
SUBMIT_LOGIN_URL = "https://zhjwxk.cic.tsinghua.edu.cn/j_acegi_formlogin_xsxk.do"
P_XNXQ = "2024-2025-2"

cookies = {"JSESSIONID": JSESSIONID}
all_data = {}

ocr = ddddocr.DdddOcr(show_ad=False)


def get_captcha() -> str:
    # 人工智能识别验证码
    response = requests.get(
        CAPTCHA_URL,
        cookies=cookies,
    )

    img = response.content
    ocr.set_ranges(5)
    result = ocr.classification(img, probability=True)
    txt = ""
    for i in result['probability']:
        txt += result['charsets'][i.index(max(i))]

    return txt


def login(account):
    captcha = get_captcha()
    data = {
        "j_username": account["username"],
        "j_password": account["password"],
        "captchaflag": "login1",
        "_login_image_": captcha,
    }
    _response = requests.post(SUBMIT_LOGIN_URL, cookies=cookies, data=data)


def get_selection_array(raw_selection_array):
    # 处理任选字段
    selection = raw_selection_array.strip("")
    if selection.startswith("("):
        selection = selection[1:].replace(")", ",")
    else:
        selection = "0," + selection
    selection_list = [int(x) for x in selection.split(",")]
    return {
        "优先志愿": selection_list[0],
        "第一志愿": selection_list[1],
        "第二志愿": selection_list[2],
        "第三志愿": selection_list[3],
    }


def get_page_courses(page_num):
    url = f"{BASE_URL_1}?m=tbzySearchBR&page={page_num}&p_xnxq={P_XNXQ}"
    response = requests.get(url, cookies=cookies)

    if response.status_code == 200:
        # 使用 BeautifulSoup 解析 HTML 内容
        soup = BeautifulSoup(response.text, "html.parser")

        # 使用正则表达式查找包含 gridData 定义的 <script> 标签
        script_tags = soup.find_all(
            "script", string=lambda text: "var gridData" in str(text)
        )

        if script_tags:
            # 遍历所有 <script> 标签,提取 gridData 定义
            for script_tag in script_tags:
                script_text = script_tag.string

                if script_text:
                    grid_data_definition = re.search(
                        r"var gridData = (.*);", script_text, re.DOTALL
                    )

                    if grid_data_definition:
                        grid_data_content = grid_data_definition.group(1).strip()

                        data = json.loads(grid_data_content)

                        for i, row in enumerate(data):
                            num = row[0]
                            idx = row[1]
                            course = {
                                "课程号": row[0],
                                "课序号": row[1],
                                "课程名": row[2],
                                "开课系": row[3],
                                "可选容量": -1 if int(row[4]) == 0 else int(row[4]),
                                "报名总人数": int(row[5]),
                                "必修报名人数": get_selection_array(row[6]),
                                "限选报名人数": get_selection_array(row[7]),
                                "任选报名人数": get_selection_array(row[8]),
                                "主讲教师": "",
                                "上课时间": "",
                            }
                            if num not in all_data:
                                all_data[num] = {}
                            all_data[num][idx] = course
        else:
            print("未找到 gridData 的定义")
    else:
        print(f"Error: {response.status_code} - {response.text}")


def get_page_pe(page_num):
    url = f"{BASE_URL_1}?m=tbzySearchTy&page={page_num}&p_xnxq={P_XNXQ}"
    response = requests.get(url, cookies=cookies)

    if response.status_code == 200:
        # 使用 BeautifulSoup 解析 HTML 内容
        soup = BeautifulSoup(response.text, "html.parser")

        # 使用正则表达式查找包含 gridData 定义的 <script> 标签
        script_tags = soup.find_all(
            "script", string=lambda text: "var gridData" in str(text)
        )

        if script_tags:
            # 遍历所有 <script> 标签,提取 gridData 定义
            for script_tag in script_tags:
                script_text = script_tag.string

                if script_text:
                    grid_data_definition = re.search(
                        r"var gridData = (.*);", script_text, re.DOTALL
                    )

                    if grid_data_definition:
                        grid_data_content = grid_data_definition.group(1).strip()

                        data = json.loads(grid_data_content)

                        for i, row in enumerate(data):
                            num = row[0]
                            idx = row[1]
                            course = {
                                "课程号": row[0],
                                "课序号": row[1],
                                "课程名": row[2],
                                "开课系": "体育部",
                                "可选容量": -1 if int(row[3]) == 0 else int(row[3]),
                                "报名总人数": int(row[4]),
                                "必修报名人数": get_selection_array(row[5]),
                                "限选报名人数": get_selection_array("0, 0, 0"),
                                "任选报名人数": get_selection_array("0, 0, 0"),
                                "主讲教师": "",
                                "上课时间": "",
                            }
                            if num not in all_data:
                                all_data[num] = {}
                            all_data[num][idx] = course

        else:
            print("未找到 gridData 的定义")
    else:
        print(f"Error: {response.status_code} - {response.text}")


def get_page_teacher(page_num):
    url = f"{BASE_URL_2}?m=kkxxSearch&page={page_num}&p_xnxq={P_XNXQ}"
    response = requests.get(url, cookies=cookies)
    if response.status_code == 200:
        # 使用 BeautifulSoup 解析 HTML 内容
        soup = BeautifulSoup(response.text, "html.parser")
        for row in soup.find_all("tr", class_="trr2"):
            cols = row.find_all("td")
            if len(cols) == 19:
                num = cols[1].text.strip()
                idx = cols[2].text.strip()
                if num in all_data:
                    if idx in all_data[num]:
                        if cols[5].find("a") is not None:
                            all_data[num][idx]["主讲教师"] = (
                                cols[5].find("a").text.strip()
                            )
                        all_data[num][idx]["上课时间"] = cols[10].text.strip()
            else:
                print(cols)
    else:
        print(f"Error: {response.status_code} - {response.text}")


if __name__ == "__main__":
    with open("account.json") as f:
        data = json.load(f)
        account = {"username": data["username"], "password": data["password"]}
        login(account)
    for page in tqdm(range(1, 158)):
        get_page_courses(page)
    for page in tqdm(range(1, 20)):
        get_page_pe(page)
    for page in tqdm(range(1, 254)):
        get_page_teacher(page)
    with open("courses.json", "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=4)
