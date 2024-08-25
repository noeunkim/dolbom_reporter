import argparse
from sheet_generator import GenerateCustomSheet

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate custom sheet")
    parser.add_argument("--child_name", default="김철수", help="아이의 이름")
    parser.add_argument("--city", default="서울시", help="도시 이름")
    parser.add_argument("--teacher_name", default="김영희", help="선생님 이름")

    args = parser.parse_args()

    custom_sheet = GenerateCustomSheet()
    custom_sheet.run(child_name=args.child_name, city=args.city, teacher_name=args.teacher_name)
