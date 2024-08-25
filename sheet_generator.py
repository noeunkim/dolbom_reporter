import os
import math
import gspread
import pandas as pd
from loguru import logger
from datetime import datetime
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials

from lxml import etree, html


class GenerateCustomSheet:
    def __init__(self):
        logger.info("Load the sheet...")

        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]

        json_key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config', 'sheet_viewer.json')
        credential = ServiceAccountCredentials.from_json_keyfile_name(json_key_path, scope)
        gc = gspread.authorize(credential)

        load_dotenv()
        google_sheet_id = os.environ.get('GOOGLE_SHEET_ID')
        doc = gc.open_by_url(f"https://docs.google.com/spreadsheets/d/{google_sheet_id}")
        self.sheet = doc.worksheet("base")

        logger.info("Load the Source HTML...")
        self.tree = html.parse(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'source', 'doc.html'))

    def _copy_pages(self, rows: int):
        added_page_num = math.ceil(rows / 5) - 1
        page_idx = 1

        def copy_elements(elem):
            new_elem = etree.Element(elem.tag)
            new_elem.text = elem.text
            new_elem.tail = elem.tail
            new_elem.attrib.update(elem.attrib)
            for child in elem:
                new_elem.append(copy_elements(child))
            return new_elem

        root = self.tree.getroot()
        body = root.find(".//body")

        for i in range(added_page_num + 1):
            if i != 0:
                if body is not None:
                    original_div = body.find(f".//div[@id='page_{page_idx}']")

                    if original_div is not None:
                        page_idx += 1
                        new_div = copy_elements(original_div)
                        new_div.set("id", f"page_{page_idx}")
                        original_div.addnext(new_div)
                else:
                    logger.error("Body element not found in the base HTML.")

        logger.info(f"* {page_idx - 1} pages are added.")

    @staticmethod
    def _format_time(time):
        hours = int(time)
        minutes = (int(time) - hours) * 60
        return f"{hours:02d}:{minutes:02d}"

    def run(self, child_name: str, city: str, teacher_name: str):
        df = pd.DataFrame(self.sheet.get_all_values())
        df = df[1:]  # remove header
        logger.info(f"Check the contents\n{df.head()}\n --- total row is {df.shape[0]}")

        # configs
        weekdays = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
        locations = {
            "이용자가정": 18,
            "돌보미가정": 22,
            "치료센터": 24,
            "이동동반": 26,
            "일상생활": 28,
            "외출/산책": 30,
            "신변처리": 32,
            "학습/놀이": 34,
        }

        # calculate
        df['day'] = pd.to_datetime(df[0])
        df[1] = df[1].astype(int)
        df[2] = df[2].astype(int)
        df['work_hours'] = df[2] - df[1]
        total_visit = df.shape[0]
        total_days = df['day'].nunique()
        total_hours = df['work_hours'].sum()
        work_months = df['day'].dt.month.unique()

        # copy the pages from base html
        self._copy_pages(df.shape[0])

        row_idx_weight = 20
        total_work_time = 0
        for index, row in df.iterrows():
            """
            Row Columns
            날짜, 시작 시간, 종료 시간, 장소, 내용
            """

            index_weight = (index - 1) % 5 + 1
            page_id = f"page_{(index - 1) // 5 + 1}"

            # 0. set the header
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[1]/div/div/div[2]/span"
            )[0]
            element.text = f"({work_months[0]})월,\u00A0({total_visit})회,\u00A0({total_days})일,({total_hours})시간"
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[3]/div/div/div/span"
            )[0]
            element.text = city
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[5]/div/div/div/span"
            )[0]
            element.set('text-align', 'right')
            element.text = child_name
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[7]/div/div/div/span"
            )[0]
            element.text = f"{teacher_name}\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0(서명)"

            # 1. set the date (row[0])
            row_weight = row_idx_weight * (index_weight - 1)
            date_object = datetime.strptime(row[0], "%Y-%m-%d")
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[{15 + row_weight}]/div/div/div[1]/span"
            )[0]
            element.text = f"{date_object.month}월 {date_object.day}일"
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[{15 + row_weight}]/div/div/div[3]/span"
            )[0]
            element.text = f"{weekdays[date_object.weekday()]}"

            # 2. set the time (row[2]-row[1])
            work_time = row[2] - row[1]
            total_work_time += work_time
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[{16 + row_weight}]/div/div/div[1]/span[1]"
            )[0]
            element.text = f"{work_time}\u00A0"
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[{16 + row_weight}]/div/div/div[2]/span"
            )[0]
            element.text = f"\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0({self._format_time(row[1])}\u00A0~\u00A0"
            element = self.tree.xpath(
                f"//*[@id=\"{page_id}\"]/div/div[2]/div[{16 + row_weight}]/div/div/div[3]/span"
            )[0]
            element.text = f"{self._format_time(row[2])})"

            # 3. set the locations (row[3])
            for idx in locations.values():
                # init locations
                element = self.tree.xpath(f"//*[@id=\"{page_id}\"]/div/div[2]/div[{idx + row_weight}]/div/div/div")[0]
                element.text = ""
            location = locations.get(row[3], None)
            if location:
                element = self.tree.xpath(f"//*[@id=\"{page_id}\"]/div/div[2]/div[{location + row_weight}]/div/div/div")[0]
                element.set("style", "padding-left: 5px; padding-bottom: 5px;")
                element.text = "O"
            else:
                logger.warning(f"Wrong location! - {index} Row: {row[3]}")

            # 4. set the details (row[4])
            element = self.tree.xpath(f"//*[@id=\"{page_id}\"]/div/div[2]/div[{19 + row_weight}]/div/div/div/span")
            if element:
                element[0].set('style', 'white-space: pre-wrap; text-align: center;')
                element[0].text = str(row[4])
            else:
                logger.error(f"Wrong detail elements for {page_id}! - {str(row[4])}")

            html.tostring(self.tree, encoding="unicode", method="html")

        self.tree.write(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output', 'doc.html'),
            encoding="utf-8",
            method="html",
        )
