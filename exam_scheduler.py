import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

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

        tk.Label(frm, text="시작일 (YYYY-MM-DD)").grid(row=0, column=0)
        tk.Label(frm, text="종료일 (YYYY-MM-DD)").grid(row=1, column=0)

        self.start_entry = tk.Entry(frm)
        self.end_entry = tk.Entry(frm)
        self.start_entry.grid(row=0, column=1)
        self.end_entry.grid(row=1, column=1)

        tk.Button(frm, text="기간 설정", command=self.set_period).grid(row=2, column=0, columnspan=2, pady=5)

        self.dates_container = tk.Frame(frm)
        self.dates_container.grid(row=3, column=0, columnspan=2, pady=10)

    def set_period(self):
        start = self.start_entry.get().strip()
        end = self.end_entry.get().strip()
        try:
            start_date = datetime.strptime(start, "%Y-%m-%d").date()
            end_date = datetime.strptime(end, "%Y-%m-%d").date()
            if start_date.year != 2025 or end_date.year != 2025:
                raise ValueError("2025년 날짜만 허용됩니다")
            if end_date < start_date:
                raise ValueError("종료일이 시작일보다 빠릅니다")
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return

        # clear previous
        for child in self.dates_container.winfo_children():
            child.destroy()
        self.exam_dates.clear()

        delta = (end_date - start_date).days
        for i in range(delta + 1):
            d = start_date + timedelta(days=i)
            date_str = d.isoformat()
            tk.Label(self.dates_container, text=date_str).grid(row=i, column=0)
            spin = tk.Spinbox(self.dates_container, from_=1, to=4, width=5)
            spin.grid(row=i, column=1)
            self.exam_dates[date_str] = spin

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

    # ----- 스케줄 생성 -----
    def build_schedule_frame(self):
        frm = self.schedule_frame
        tk.Button(frm, text="스케줄 생성", command=self.generate_schedule).pack(pady=5)
        self.schedule_text = tk.Text(frm, width=80, height=25)
        self.schedule_text.pack(pady=10)

    def generate_schedule(self):
        # finalize exam dates with periods
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
            t.assignments = {"main":0, "assistant":0, "selfstudy":0}

        schedule_lines = []
        subj_iter = iter(self.subjects)
        for date, period in dates:
            try:
                subject = next(subj_iter)
            except StopIteration:
                subject = None
            subject_name = subject.name if subject else "자율학습"
            forbidden = subject.teacher if subject else None

            available = [t for t in self.teachers.values() if (date, period) not in t.unavailable and t.name != forbidden]
            if len(available) < 3:
                messagebox.showerror("오류", f"{date} {period}교시에 배정 가능한 교사가 부족합니다")
                return

            # simple round-robin using assignment counts
            available.sort(key=lambda t: (t.assignments["main"], t.assignments["assistant"], t.assignments["selfstudy"]))
            main = available[0]
            assistant = available[1]
            selfstudy = available[2]
            main.assignments["main"] += 1
            assistant.assignments["assistant"] += 1
            selfstudy.assignments["selfstudy"] += 1
            line = f"{date} {period}교시 - {subject_name} : 정감독 {main.name}, 부감독 {assistant.name}, 자율학습 {selfstudy.name}"
            schedule_lines.append(line)

        self.schedule_text.delete("1.0", tk.END)
        self.schedule_text.insert(tk.END, "\n".join(schedule_lines))

        # totals
        self.schedule_text.insert(tk.END, "\n\n[교사별 배정 합계]\n")
        for t in self.teachers.values():
            tot_line = f"{t.name}: 정감독 {t.assignments['main']}, 부감독 {t.assignments['assistant']}, 자율학습 {t.assignments['selfstudy']}"
            self.schedule_text.insert(tk.END, tot_line + "\n")

if __name__ == "__main__":
    app = ExamScheduler()
    app.mainloop()
