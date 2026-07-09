import fitz
import os
import re
import pandas as pd

df = pd.read_excel("word-data-test-testing.xlsx")

folder_path = "../CAS-Original-PDFs"

restricted_contents = [
    "\"", "-", ".", "°", "/", "\\", ",", "(", ")"
]


arr = []

for name in os.listdir(folder_path):

    full_path = os.path.join(folder_path, name)
    dr_name = name.split(".")[0]

    if os.path.isfile(full_path):

        file_df = df.loc[df["P&ID DRG NO."] == dr_name]

        Hashmap = file_df.to_dict(orient="records")

        doc = fitz.open(full_path)

        page = doc[0]

        page_coordinates = page.rect

        width = page_coordinates.width
        height = page_coordinates.height

        print(dr_name)

        count = 1

        drawn_rects = []

        for row in Hashmap:

            X1 = float(row["X1"])
            Y1 = float(row["Y1"])
            X2 = float(row["X2"])
            Y2 = float(row["Y2"])

            custom_rect = fitz.Rect(X1, Y1 - 18, X2, Y2)
            rect = fitz.Rect(X1, Y1, X2, Y2)

            raw_text = page.get_text("text", clip=rect)

            lines = [
                line.strip()
                for line in raw_text.splitlines()
                if line.strip()
            ]

            # Keep only tags such as:
            # 04PR\n087A
            # PSV\n101
            # 22P\n44CV

            if len(lines) != 2:
                continue

            if not all(
                re.fullmatch(r"[A-Z0-9]+", line)
                for line in lines
            ):
                continue

            text = "".join(lines)

            if (
                any(char in text for char in restricted_contents)
                or "NOTE" in text.upper()
            ):
                continue

            overlap = False
            index = 0
            existing_text = ""

            for existing in drawn_rects:

                if rect.intersects(existing):

                    overlap = True
                    index = drawn_rects.index(existing)

                    existing_raw = page.get_text(
                        "text",
                        clip=existing
                    )

                    existing_text = existing_raw.replace(
                        " ", ""
                    ).replace(
                        "\n", ""
                    )

                    break

            data = {
                "P&ID DRG NO.": dr_name,
                "BOX ID": count,
                "TAG ID": text,
                "TAG TYPE": "Instrument",
                "X1": X1,
                "Y1": Y1,
                "X2": X2,
                "Y2": Y2,
                "HEIGHT": height,
                "WIDTH": width
            }

            if overlap and len(text) < len(existing_text):
                continue

            elif overlap and len(text) > len(existing_text):

                ex1, ey1, ex2, ey2 = drawn_rects[index]

                drawn_rects[index] = rect

                for i, d in enumerate(arr):

                    if (
                        d["X1"] == ex1 and
                        d["Y1"] == ey1 and
                        d["X2"] == ex2 and
                        d["Y2"] == ey2
                    ):
                        print(d["BOX ID"], "removed")
                        arr.pop(i)
                        break

                arr.append(data)

                print(data)

                page.draw_rect(
                    rect,
                    color=(1, 0, 0),
                    width=1,
                    overlay=True
                )

            else:

                drawn_rects.append(rect)

                arr.append(data)

                page.draw_rect(
                    rect,
                    color=(0, 1, 0),
                    width=1,
                    overlay=True
                )

            page.insert_textbox(
                custom_rect,
                f"B-{count}",
                fontsize=8,
                color=(0, 0, 1),
                overlay=True,
                align=1
            )

            count += 1

        doc.save(f"copilot-7/{name}")
        doc.close()

df_backplot = pd.DataFrame(arr)

df_backplot.to_excel(
    "7-Files-new-testing-copilot.xlsx",
    sheet_name="sheet1",
    index=False
)
