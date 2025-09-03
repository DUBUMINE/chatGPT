import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar
from datetime import datetime, timedelta, date
import json
from openpyxl import Workbook

class Teacher:
    def __init__(self, name):
        self.name = name
        self.unavailable = set()  # set of (date_str, period)
        self.assignments = {"main": 0, "assistant": 0, "selfstudy": 0}

    def add_unavailability(self, date_str, period):
        self.unavailable.add((date_str, period))

class Subject:
    def __init__(self, name, teacher):
        self.name = name
        self.teacher = teacher

class ExamScheduler(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("시험 감독 스케줄러")
        self.geometry("900x600")

        self.exam_dates = {}  # date_str -> periods
        self.subjects = []
        self.teachers = {}  # name -> Teacher
        self.grade_count = 1
        self.class_count = 1
        self.generated_schedule = []

        self.create_widgets()

    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.period_frame = ttk.Frame(notebook)
        notebook.add(self.period_frame, text="시험 기간")
        self.subject_frame = ttk.Frame(notebook)
        notebook.add(self.subject_frame, text="과목")
        self.teacher_frame = ttk.Frame(notebook)
        notebook.add(self.teacher_frame, text="교사")
        self.schedule_frame = ttk.Frame(notebook)
        notebook.add(self.schedule_frame, text="스케줄")

        self.build_period_frame()
        self.build_subject_frame()
        self.build_teacher_frame()
        self.build_schedule_frame()

    # ----- 시험 기간 설정 -----
    def build_period_frame(self):
        frm = self.period_frame

        tk.Label(frm, text="학년 수").grid(row=0, column=0)
        tk.Label(frm, text="학급 수").grid(row=0, column=2)
        self.grade_spin = tk.Spinbox(frm, from_=1, to=12, width=5)
        self.class_spin = tk.Spinbox(frm, from_=1, to=20, width=5)
        self.grade_spin.grid(row=0, column=1)
        self.class_spin.grid(row=0, column=3)

        self.start_cal = Calendar(
            frm,
            selectmode="day",
            year=2025,
            month=1,
            day=1,
            mindate=date(2025, 1, 1),
            maxdate=date(2030, 12, 31),
            firstweekday="sunday",
        )
        self.start_cal.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        self.end_cal = Calendar(
            frm,
            selectmode="day",
            year=2025,
            month=1,
            day=1,
            mindate=date(2025, 1, 1),
            maxdate=date(2030, 12, 31),
            firstweekday="sunday",
        )
        self.end_cal.grid(row=1, column=2, columnspan=2, padx=5, pady=5)

        tk.Button(frm, text="기간 설정", command=self.set_period).grid(
            row=2, column=0, columnspan=4, pady=5
        )

        self.dates_container = tk.Frame(frm)
        self.dates_container.grid(row=3, column=0, columnspan=4, pady=10)

    def set_period(self):
        start_date = self.start_cal.selection_get()
        end_date = self.end_cal.selection_get()
        if end_date < start_date:
            messagebox.showerror("오류", "종료일이 시작일보다 빠릅니다")
            return
        self.grade_count = int(self.grade_spin.get())
        self.class_count = int(self.class_spin.get())

        for child in self.dates_container.winfo_children():
            child.destroy()
        self.exam_dates.clear()

        delta = (end_date - start_date).days
        row = 0
        for i in range(delta + 1):
            d = start_date + timedelta(days=i)
            if d.weekday() >= 5:
                continue  # skip weekends
            date_str = d.isoformat()
            tk.Label(self.dates_container, text=date_str).grid(row=row, column=0)
            spin = tk.Spinbox(self.dates_container, from_=1, to=4, width=5)
            spin.grid(row=row, column=1)
            self.exam_dates[date_str] = spin
            row += 1

    # ----- 과목 -----
    def build_subject_frame(self):
        frm = self.subject_frame
        tk.Label(frm, text="과목명").grid(row=0, column=0)
        tk.Label(frm, text="담당교사").grid(row=0, column=1)
        self.subject_name = tk.Entry(frm)
        self.subject_teacher = tk.Entry(frm)
        self.subject_name.grid(row=1, column=0)
        self.subject_teacher.grid(row=1, column=1)
        tk.Button(frm, text="추가", command=self.add_subject).grid(row=1, column=2, padx=5)

        self.subject_list = tk.Listbox(frm, width=60)
        self.subject_list.grid(row=2, column=0, columnspan=3, pady=10)

        tk.Button(frm, text="저장", command=self.save_subjects).grid(row=3, column=0, pady=5)
        tk.Button(frm, text="로드", command=self.load_subjects).grid(row=3, column=1, pady=5)

    def add_subject(self):
        name = self.subject_name.get().strip()
        teacher = self.subject_teacher.get().strip()
        if not name or not teacher:
            messagebox.showwarning("경고", "과목명과 담당교사를 입력하세요")
            return
        self.subjects.append(Subject(name, teacher))
        self.subject_list.insert(tk.END, f"{name} - {teacher}")
        self.subject_name.delete(0, tk.END)
        self.subject_teacher.delete(0, tk.END)

    def save_subjects(self):
        if not self.subjects:
            messagebox.showwarning("경고", "저장할 과목이 없습니다")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json")]
        )
        if path:
            data = [{"name": s.name, "teacher": s.teacher} for s in self.subjects]
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

    def load_subjects(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                messagebox.showerror("오류", str(e))
                return
            self.subjects.clear()
            self.subject_list.delete(0, tk.END)
            for item in data:
                s = Subject(item["name"], item["teacher"])
                self.subjects.append(s)
                self.subject_list.insert(tk.END, f"{s.name} - {s.teacher}")

    # ----- 교사 -----
    def build_teacher_frame(self):
        frm = self.teacher_frame
        tk.Label(frm, text="교사 이름").grid(row=0, column=0)
        self.teacher_name = tk.Entry(frm)
        self.teacher_name.grid(row=0, column=1)
        tk.Button(frm, text="교사 추가", command=self.add_teacher).grid(row=0, column=2, padx=5)

        self.teacher_list = tk.Listbox(frm, width=40)
        self.teacher_list.grid(row=1, column=0, columnspan=3, pady=10)

        tk.Button(frm, text="불참 추가", command=self.add_unavailability).grid(row=2, column=0, columnspan=3, pady=5)

        tk.Button(frm, text="저장", command=self.save_teachers).grid(row=3, column=0, pady=5)
        tk.Button(frm, text="로드", command=self.load_teachers).grid(row=3, column=1, pady=5)

    def add_teacher(self):
        name = self.teacher_name.get().strip()
        if not name:
            return
        if name in self.teachers:
            messagebox.showerror("오류", "이미 존재하는 교사")
            return
        self.teachers[name] = Teacher(name)
        self.teacher_list.insert(tk.END, name)
        self.teacher_name.delete(0, tk.END)

    def add_unavailability(self):
        selection = self.teacher_list.curselection()
        if not selection:
            messagebox.showwarning("경고", "교사를 선택하세요")
            return
        name = self.teacher_list.get(selection[0])
        teacher = self.teachers[name]

        win = tk.Toplevel(self)
        win.title(f"{name} 불가능 시간")
        tk.Label(win, text="날짜 YYYY-MM-DD").grid(row=0, column=0)
        tk.Label(win, text="교시").grid(row=1, column=0)
        date_entry = tk.Entry(win)
        period_spin = tk.Spinbox(win, from_=1, to=4, width=5)
        date_entry.grid(row=0, column=1)
        period_spin.grid(row=1, column=1)

        def save():
            date = date_entry.get().strip()
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("오류", "날짜 형식이 잘못되었습니다")
                return
            period = int(period_spin.get())
            teacher.add_unavailability(date, period)
            win.destroy()

        tk.Button(win, text="저장", command=save).grid(row=2, column=0, columnspan=2)

    def save_teachers(self):
        if not self.teachers:
            messagebox.showwarning("경고", "저장할 교사가 없습니다")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json")]
        )
        if path:
            data = []
            for t in self.teachers.values():
                data.append(
                    {
                        "name": t.name,
                        "unavailable": [
                            {"date": d, "period": p} for d, p in t.unavailable
                        ],
                    }
                )
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

    def load_teachers(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception as e:
                messagebox.showerror("오류", str(e))
                return
            self.teachers.clear()
            self.teacher_list.delete(0, tk.END)
            for item in data:
                t = Teacher(item["name"])
                for u in item.get("unavailable", []):
                    t.add_unavailability(u["date"], u["period"])
                self.teachers[t.name] = t
                self.teacher_list.insert(tk.END, t.name)

    # ----- 스케줄 생성 -----
    def build_schedule_frame(self):
        frm = self.schedule_frame
        tk.Button(frm, text="스케줄 생성", command=self.generate_schedule).pack(pady=5)
        tk.Button(frm, text="엑셀 저장", command=self.export_excel).pack(pady=5)
        self.schedule_text = tk.Text(frm, width=80, height=25)
        self.schedule_text.pack(pady=10)

    def generate_schedule(self):
        dates = []
        for date_str, spin in self.exam_dates.items():
            periods = int(spin.get())
            self.exam_dates[date_str] = periods
            for p in range(1, periods + 1):
                dates.append((date_str, p))

        if not dates:
            messagebox.showerror("오류", "시험 기간이 설정되지 않았습니다")
            return
        if not self.subjects:
            messagebox.showerror("오류", "과목이 없습니다")
            return
        if len(self.teachers) < 3:
            messagebox.showerror("오류", "교사가 최소 3명 필요합니다")
            return

        for t in self.teachers.values():
            t.assignments = {"main": 0, "assistant": 0, "selfstudy": 0}

        self.generated_schedule = []
        schedule_lines = []
        subj_iter = iter(self.subjects)
        for date, period in dates:
            try:
                subject = next(subj_iter)
            except StopIteration:
                subject = None
            subject_name = subject.name if subject else "자율학습"
            forbidden = subject.teacher if subject else None

            available = [
                t
                for t in self.teachers.values()
                if (date, period) not in t.unavailable and t.name != forbidden
            ]
            if len(available) < 3:
                messagebox.showerror(
                    "오류", f"{date} {period}교시에 배정 가능한 교사가 부족합니다"
                )
                return

            available.sort(
                key=lambda t: (
                    t.assignments["main"],
                    t.assignments["assistant"],
                    t.assignments["selfstudy"],
                )
            )
            main = available[0]
            assistant = available[1]
            selfstudy = available[2]
            main.assignments["main"] += 1
            assistant.assignments["assistant"] += 1
            selfstudy.assignments["selfstudy"] += 1
            line = (
                f"{date} {period}교시 - {subject_name} : 정감독 {main.name}, 부감독 {assistant.name}, 자율학습 {selfstudy.name}"
            )
            schedule_lines.append(line)
            self.generated_schedule.append(
                (date, period, subject_name, main.name, assistant.name, selfstudy.name)
            )

        self.schedule_text.delete("1.0", tk.END)
        self.schedule_text.insert(tk.END, "\n".join(schedule_lines))

        self.schedule_text.insert(tk.END, "\n\n[교사별 배정 합계]\n")
        for t in self.teachers.values():
            tot_line = (
                f"{t.name}: 정감독 {t.assignments['main']}, 부감독 {t.assignments['assistant']}, 자율학습 {t.assignments['selfstudy']}"
            )
            self.schedule_text.insert(tk.END, tot_line + "\n")

    def export_excel(self):
        if not self.generated_schedule:
            messagebox.showwarning("경고", "스케줄이 없습니다")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Schedule"
        ws.append(["날짜", "교시", "과목", "정감독", "부감독", "자율학습"])
        for row in self.generated_schedule:
            ws.append(list(row))

        ws2 = wb.create_sheet("Summary")
        ws2.append(["교사", "정감독", "부감독", "자율학습"])
        for t in self.teachers.values():
            ws2.append(
                [
                    t.name,
                    t.assignments["main"],
                    t.assignments["assistant"],
                    t.assignments["selfstudy"],
                ]
            )

        ws3 = wb.create_sheet("Info")
        ws3.append(["학년 수", self.grade_count])
        ws3.append(["학급 수", self.class_count])
        wb.save(path)

if __name__ == "__main__":
    app = ExamScheduler()
    app.mainloop()
