# Dev by hzuani | 2026.02.25 Updated (엑셀 셀 병합 완벽 대응 + 통합 PDF)
import pandas as pd
import win32com.client as win32
import os
import re
import sys
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfWriter

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ABSENCE_HWP = os.path.join(BASE_DIR, "2026년 결석 신고서.hwp")
RECOGNIZED_HWP = os.path.join(BASE_DIR, "2026년 인정 출결 신고서.hwp")
OUTPUT_DIR = os.path.join(BASE_DIR, "출력물_완성")

def select_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="나이스 엑셀 파일을 선택하세요",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

def extract_date(date_str):
    try:
        nums = re.findall(r'\d+', str(date_str))
        if len(nums) >= 3:
            return datetime(int(nums[0]), int(nums[1]), int(nums[2]))
    except:
        return None
    return None

def group_by_date(r_list, gap_days=7):
    """날짜 기준으로 연속된 기록을 그룹화 (gap_days 이내면 같은 그룹)"""
    r_list = sorted(r_list, key=lambda x: x['date'])
    groups = []
    cur = [r_list[0]]
    for i in range(1, len(r_list)):
        if (r_list[i]['date'] - r_list[i-1]['date']).days <= gap_days:
            cur.append(r_list[i])
        else:
            groups.append(cur)
            cur = [r_list[i]]
    groups.append(cur)
    return groups

def get_reason(group):
    """그룹에서 사유 추출 — 마지막으로 채워진 사유를 반환"""
    reason = ""
    for d in group:
        if d['reason'] and str(d['reason']).strip():
            reason = str(d['reason']).strip()
    return reason

def process_attendance(excel_file):
    # 파일명에서 학년/반 자동 추출 (예: "2학년 2반 월별 출결 현황.xlsx")
    fname = os.path.basename(excel_file)
    m = re.search(r'(\d+)학년\s*(\d+)반', fname)
    if m:
        GRADE, CLASS = m.group(1), m.group(2)
    else:
        print(f"⚠️ 파일명에서 학년/반을 찾지 못했습니다. 파일명: {fname}")
        GRADE, CLASS = "?", "?"

    print(f"📂 파일 읽는 중: {excel_file} ({GRADE}학년 {CLASS}반)")

    try:
        df = pd.read_excel(excel_file)
        header_idx = next((i for i, row in df.iterrows() if "성명" in row.values), 0)
        df = pd.read_excel(excel_file, header=header_idx).dropna(subset=['일자'])

        # 셀 병합 해결 — 빈칸을 같은 학생 데이터로 채워줌
        df['번호'] = df['번호'].ffill()
        df['성명'] = df['성명'].ffill()
        df['출결구분'] = df.groupby(['번호', '성명'])['출결구분'].ffill()
        df['사유'] = df.groupby(['번호', '성명'])['사유'].ffill()

    except Exception as e:
        return f"엑셀 오류: {str(e)}"

    students = {}
    for _, row in df.iterrows():
        if pd.isna(row['출결구분']):
            continue
        try:
            s_no = str(int(float(row['번호']))) if not pd.isna(row['번호']) else "?"
        except (ValueError, TypeError):
            s_no = str(row['번호'])
        s_name = str(row['성명']).strip()
        key = f"{s_no}_{s_name}"

        # 날짜 파싱 실패한 행은 건너뜀
        d = extract_date(row['일자'])
        if d is None:
            continue

        if key not in students:
            students[key] = []
        students[key].append({
            'date': d,
            'type': str(row['출결구분']).strip(),
            'reason': str(row['사유']).strip() if not pd.isna(row['사유']) else "",
            'period': str(row['결시교시']) if not pd.isna(row['결시교시']) else ""
        })

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True
    except Exception as e:
        return f"한글 실행 실패: {str(e)}"

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    count_files = 0
    pdf_list = []

    try:
        for key, records in students.items():
            # 이름에 _ 포함되어도 안전하게 분리
            s_no, s_name = key.split('_', 1)

            # [1] 일반 결석
            normal_absences = [r for r in records if "결석" in r['type'] and ("미인정" in r['type'] or "인정" not in r['type'])]
            grouped_dict = {}
            for r in normal_absences:
                k = r['type']
                if k not in grouped_dict:
                    grouped_dict[k] = []
                grouped_dict[k].append(r)

            for _, r_list in grouped_dict.items():
                for group in group_by_date(r_list):
                    first, last = group[0], group[-1]
                    t_type = first['type']
                    g_reason = get_reason(group)

                    hwp.Open(ABSENCE_HWP)
                    hwp.PutFieldText("grade", GRADE); hwp.PutFieldText("class", CLASS)
                    hwp.PutFieldText("number", s_no); hwp.PutFieldText("name", s_name)

                    s_dt, e_dt = first['date'], last['date']
                    r_dt = last['date'] + timedelta(days=1)

                    hwp.PutFieldText("sy", str(s_dt.year)[-2:]); hwp.PutFieldText("sm", str(s_dt.month)); hwp.PutFieldText("sd", str(s_dt.day))
                    hwp.PutFieldText("ey", str(e_dt.year)[-2:]); hwp.PutFieldText("em", str(e_dt.month)); hwp.PutFieldText("ed", str(e_dt.day))
                    hwp.PutFieldText("days", str(len(group))); hwp.PutFieldText("reason", g_reason)

                    if (e_dt - s_dt).days + 1 != len(group):
                        hwp.PutFieldText("sub_dates", ", ".join([f"{d['date'].month}/{d['date'].day}" for d in group]))

                    hwp.PutFieldText("ty", str(r_dt.year)[-2:]); hwp.PutFieldText("tm", str(r_dt.month)); hwp.PutFieldText("td", str(r_dt.day))
                    hwp.PutFieldText("cm", str(s_dt.month)); hwp.PutFieldText("cd", str(s_dt.day)); hwp.PutFieldText("ch", "8")

                    hwp.PutFieldText("chk_disease", ""); hwp.PutFieldText("chk_unauth", ""); hwp.PutFieldText("chk_other", "")
                    if "질병" in t_type: hwp.PutFieldText("chk_disease", "V")
                    elif "미인정" in t_type: hwp.PutFieldText("chk_unauth", "V")
                    else: hwp.PutFieldText("chk_other", "V")

                    fname_hwp = os.path.join(OUTPUT_DIR, f"{s_no}번_{s_name}_결석_{s_dt.strftime('%m%d')}.hwp")
                    hwp.SaveAs(fname_hwp)

                    pdf_name = fname_hwp.replace(".hwp", ".pdf")
                    hwp.SaveAs(pdf_name, "PDF")
                    pdf_list.append(pdf_name)
                    count_files += 1

            # [2] 인정 출결
            recognized_records = [r for r in records if "인정" in r['type'] and "미인정" not in r['type']]

            rec_absences = [r for r in recognized_records if "결석" in r['type']]
            grouped_rec_abs = {}
            for r in rec_absences:
                k = r['type']
                if k not in grouped_rec_abs:
                    grouped_rec_abs[k] = []
                grouped_rec_abs[k].append(r)

            for _, r_list in grouped_rec_abs.items():
                for group in group_by_date(r_list):
                    first, last = group[0], group[-1]
                    g_reason = get_reason(group)

                    hwp.Open(RECOGNIZED_HWP)
                    hwp.PutFieldText("grade", GRADE); hwp.PutFieldText("class", CLASS)
                    hwp.PutFieldText("number", s_no); hwp.PutFieldText("name", s_name)

                    s_dt, e_dt = first['date'], last['date']
                    r_dt = last['date'] + timedelta(days=1)

                    hwp.PutFieldText("sy", str(s_dt.year)[-2:]); hwp.PutFieldText("sm", str(s_dt.month)); hwp.PutFieldText("sd", str(s_dt.day))
                    hwp.PutFieldText("ey", str(e_dt.year)[-2:]); hwp.PutFieldText("em", str(e_dt.month)); hwp.PutFieldText("ed", str(e_dt.day))
                    hwp.PutFieldText("days", str(len(group))); hwp.PutFieldText("reason", g_reason)

                    if (e_dt - s_dt).days + 1 != len(group):
                        hwp.PutFieldText("sub_dates", ", ".join([f"{d['date'].month}/{d['date'].day}" for d in group]))

                    hwp.PutFieldText("dy", ""); hwp.PutFieldText("dm", ""); hwp.PutFieldText("dd", "")
                    hwp.PutFieldText("per_s", ""); hwp.PutFieldText("per_e", "")

                    hwp.PutFieldText("ty", str(r_dt.year)[-2:]); hwp.PutFieldText("tm", str(r_dt.month)); hwp.PutFieldText("td", str(r_dt.day))
                    hwp.PutFieldText("cm", str(s_dt.month)); hwp.PutFieldText("cd", str(s_dt.day)); hwp.PutFieldText("ch", "8")

                    hwp.PutFieldText("chk_attend_abs", "V"); hwp.PutFieldText("chk_late", ""); hwp.PutFieldText("chk_early", ""); hwp.PutFieldText("chk_result", "")
                    hwp.PutFieldText("type_txt", "결석")

                    fname_hwp = os.path.join(OUTPUT_DIR, f"{s_no}번_{s_name}_인정결석_{s_dt.strftime('%m%d')}.hwp")
                    hwp.SaveAs(fname_hwp)

                    pdf_name = fname_hwp.replace(".hwp", ".pdf")
                    hwp.SaveAs(pdf_name, "PDF")
                    pdf_list.append(pdf_name)
                    count_files += 1

            # [2-B] 인정 기타 (지각, 조퇴, 결과)
            rec_others = [r for r in recognized_records if "결석" not in r['type']]
            for item in rec_others:
                hwp.Open(RECOGNIZED_HWP)
                hwp.PutFieldText("grade", GRADE); hwp.PutFieldText("class", CLASS)
                hwp.PutFieldText("number", s_no); hwp.PutFieldText("name", s_name)

                d_dt = item['date']
                r_dt = d_dt + timedelta(days=1)

                hwp.PutFieldText("sy", ""); hwp.PutFieldText("sm", ""); hwp.PutFieldText("sd", "")
                hwp.PutFieldText("ey", ""); hwp.PutFieldText("em", ""); hwp.PutFieldText("ed", "")
                hwp.PutFieldText("days", ""); hwp.PutFieldText("sub_dates", "")

                hwp.PutFieldText("dy", str(d_dt.year)[-2:]); hwp.PutFieldText("dm", str(d_dt.month)); hwp.PutFieldText("dd", str(d_dt.day))
                p_list = [p.strip() for p in item['period'].replace('"', '').split(',') if p.strip()]
                hwp.PutFieldText("per_s", p_list[0].replace("교시", "") if p_list else "")
                hwp.PutFieldText("per_e", p_list[-1].replace("교시", "") if p_list else "")

                hwp.PutFieldText("reason", item['reason'])
                hwp.PutFieldText("ty", str(r_dt.year)[-2:]); hwp.PutFieldText("tm", str(r_dt.month)); hwp.PutFieldText("td", str(r_dt.day))
                hwp.PutFieldText("cm", str(d_dt.month)); hwp.PutFieldText("cd", str(d_dt.day)); hwp.PutFieldText("ch", "8")

                hwp.PutFieldText("chk_attend_abs", ""); hwp.PutFieldText("chk_late", ""); hwp.PutFieldText("chk_early", ""); hwp.PutFieldText("chk_result", "")
                txt = ""
                if "지각" in item['type']:
                    hwp.PutFieldText("chk_late", "V")
                    txt = "지각"
                elif "조퇴" in item['type']:
                    hwp.PutFieldText("chk_early", "V")
                    txt = "조퇴"
                elif "결과" in item['type']:
                    hwp.PutFieldText("chk_result", "V")
                    txt = "결과"

                hwp.PutFieldText("type_txt", txt)

                fname_hwp = os.path.join(OUTPUT_DIR, f"{s_no}번_{s_name}_{item['type']}_{d_dt.strftime('%m%d')}.hwp")
                hwp.SaveAs(fname_hwp)

                pdf_name = fname_hwp.replace(".hwp", ".pdf")
                hwp.SaveAs(pdf_name, "PDF")
                pdf_list.append(pdf_name)
                count_files += 1

    finally:
        # 예외가 발생해도 한글 프로세스를 반드시 종료
        hwp.Quit()

    if pdf_list:
        merger = PdfWriter()
        for pdf in pdf_list:
            merger.append(pdf)

        merged_pdf_path = os.path.join(OUTPUT_DIR, "0_통합본_출결서류.pdf")
        merger.write(merged_pdf_path)
        merger.close()

        for pdf in pdf_list:
            try:
                os.remove(pdf)
            except FileNotFoundError:
                pass

    return (
        f"작업 완료!\n"
        f"총 {count_files}건의 서류가 생성되었습니다.\n\n"
        f"출력 위치: {OUTPUT_DIR}\n"
        f"통합 PDF: {os.path.join(OUTPUT_DIR, '0_통합본_출결서류.pdf')}"
    )

if __name__ == "__main__":
    excel_path = select_file()
    if excel_path:
        result_msg = process_attendance(excel_path)
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("결과", result_msg)
