import streamlit as st
import time
import pdfplumber
import re
import pandas as pd

HSC_SCIENCE_ID = 1
DIPLOMA_ID = 2

st.title("UPLOAD DIPLOMA RESULT PDF")

class PdDataFrame:
    def __init__(self):
        self.objs = {}

    def get_df(self, obj_name, header):
        if obj_name not in self.objs.keys():
            df_dict = {}
            for head in header:
                df_dict[head] = []
            self.objs[obj_name] = pd.DataFrame(df_dict)
        return self.objs[obj_name]


class DiplomaParser:
    def parser_pdf(self, _file, end=None):
        with pdfplumber.open(_file) as pdf:
            if end is None:
                end = len(pdf.pages)
            for idx in range(0, end):
                raw_data = pdf.pages[idx].extract_text()
                collegeEndIdx = raw_data.find("College") + len("College")
                SeatNoStartIdx = raw_data.find("SEAT NO")
                TotalMarksEndIdx = raw_data.find("Total Marks") + len("Total Marks") + len(":  1000")
                ResultDateStartIdx = raw_data.find("Result Date")
                titleInfo = raw_data[SeatNoStartIdx:TotalMarksEndIdx]
                HeadInfoInfo = raw_data[:collegeEndIdx]
                ClgInfo = []
                for hdIdx, hd in enumerate(HeadInfoInfo.splitlines()):
                    hd = hd.strip()
                    if hdIdx == 1:
                        sem = re.search(r'RESULT SHEET FOR THE(.*?)EXAMINATION', hd)
                        if sem:
                            sem = sem.group(1).strip()
                        ClgInfo.append(sem)
                    if hdIdx == 2:
                        instCourse = [rn.strip() for rn in hd.split("COURSE :") if len(rn) > 0]
                        CourseCode = (instCourse[1].split(" ")[0]).strip()
                        instCode = (instCourse[0].split(" ")[2]).strip()
                        ClgInfo.append(instCode)
                        ClgInfo.append(CourseCode)

                subjectInfo = raw_data[collegeEndIdx:SeatNoStartIdx]
                subjectInfo = subjectInfo.strip()
                subjectLines = subjectInfo.splitlines()
                studentMarkGp = [subjectLines[i:i + 5] for i in range(0, len(subjectLines), 5)]
                studentSubs = []
                for mark in studentMarkGp:
                    subs = mark[1].split(" ")
                    types = mark[2].split(" ")
                    final_sub = [f"{subs[i]}-{types[i]}" for i in range(len(subs))]
                    studentSubs.extend(final_sub)

                studentInfo = raw_data[TotalMarksEndIdx:ResultDateStartIdx].strip()
                lines = studentInfo.splitlines()

                bRead = False
                studentName = ""
                status = ""
                totalMarks = raw_data[TotalMarksEndIdx - len(" 1000"):TotalMarksEndIdx].strip()
                for lnNo, lnInfo in enumerate(lines):
                    if not bRead:
                        seatNo = -1
                        studentMark = []
                        seatNoCheck = [ln.strip() for ln in lnInfo.split(" ") if len(ln) > 0]
                        if len(seatNoCheck) >= 7:
                            if seatNoCheck[0].isdigit() and (
                                    seatNoCheck[5].isupper() and len(seatNoCheck[5]) == 1) and (
                                    seatNoCheck[6].isupper() and len(seatNoCheck[6]) == 1):
                                seatNo = int(seatNoCheck[0])
                        if seatNo > 0:
                            bRead = True
                            studentName = " ".join(seatNoCheck[2:5])
                            status = seatNoCheck[5].strip()
                    else:
                        if "Total" in lnInfo and "Result" in lnInfo:
                            creditsIdx = lnInfo.find("TCALSE")
                            if creditsIdx == -1:
                                resultInfo = [ln.strip() for ln in lnInfo[lnInfo.find("Total"):].split(" ") if
                                              len(ln) > 0]
                            else:
                                resultInfo = [ln.strip() for ln in lnInfo[lnInfo.find("Total"):creditsIdx].split(" ") if
                                              len(ln) > 0]
                            obtainedMark = int(resultInfo[2])
                            result = " ".join(resultInfo[5:])
                            studentMark = [studentMark[i:i + 3] for i in range(0, len(studentMark), 3)]
                            resultsMark = []
                            for mark in studentMark:
                                resultsMark.extend(mark[-1])
                            sem, instituteId, course = ClgInfo
                            yield DIPLOMA_ID, instituteId, course, sem, seatNo, studentName, obtainedMark, status, result, totalMarks, idx + 1, resultsMark, studentSubs
                            bRead = False
                            studentMark.clear()
                        else:
                            markInfo = [ln[:3] for ln in lnInfo.split(" ") if len(ln) > 0]
                            if len(markInfo) > 0:
                                studentMark.append(markInfo)

    def do_parsing(self, input_file, output_file):
        HEADER_1ST = ["ROLL NO", "NAME"]
        HEADER_LAST = ["MARKS", "STATUS", "RESULT", "TOTAL", "PAGE"]
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        try:
            df_obj = PdDataFrame()
            for data in self.parser_pdf(input_file):
                subjectList = data[-1]
                unique_name = "_".join(data[1:4]).replace(" ", "_")
                final_header = HEADER_1ST + subjectList + HEADER_LAST
                temp = df_obj.get_df(unique_name, final_header)
                final_data = list(data[4:6]) + list(data[-2]) + list(data[6:11])
                if len(final_data) == len(final_header):
                    temp.loc[len(temp.index)] = final_data
                else:
                    pass
                    #print("Mismatched columns, skipping row:", final_data)
            for key, df in df_obj.objs.items():
                startrow = 2
                heading = pd.DataFrame({key: []})
                heading.to_excel(writer, sheet_name=key, startrow=startrow)
                startrow = startrow + 1
                df.to_excel(writer, sheet_name=key, startrow=startrow, index=False)
                startrow += (df.shape[0] + 4)
                heading = pd.DataFrame({"Filter by Marks": []})
                heading.to_excel(writer, sheet_name=key, startrow=startrow)
                startrow = startrow + 1
                pec_df = df.sort_values(by=["MARKS"], ascending=False)
                pec_df.to_excel(writer, sheet_name=key, startrow=startrow, index=False)
            writer.close()
            return True
        except Exception as e:
            writer.close()
            print(e)
            return False

# Example usage:
from io import BytesIO

pdf = st.file_uploader("UPLOAD A FILE")
try:
    if pdf:
        dp = DiplomaParser()
        dp.do_parsing(pdf, "DiplomaResult.xlsx")
    
        time.sleep(5)
        data = pd.ExcelFile("DiplomaResult.xlsx")
    
        # Save the Excel file to a BytesIO object
        excel_data = BytesIO()
        with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
            for sheet_name in data.sheet_names:
                df = data.parse(sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    
        excel_data.seek(0)
    
        # Use BytesIO object as the data argument for download_button
        st.download_button("Download File", excel_data, file_name="DiplomaResult.xlsx", mime="text/csv")
expect Exception:
    st.error("PLEASE CHECK FILE FIORMAT!!!")
