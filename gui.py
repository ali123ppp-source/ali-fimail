import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import main  # نستدعي ملف main.py الذي يحتوي على الكود الرئيسي

class FamilySorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("نظام فرز وتنسيق العوائل")
        self.root.geometry("500x350")
        self.root.configure(padx=20, pady=20)
        
        self.input_file_path = None
        
        # --- تصميم الواجهة ---
        
        # 1. عنوان رئيسي
        title_label = tk.Label(root, text="نظام ترتيب بطاقات التموين والعوائل", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 2. حقل اختيار الملف
        self.file_label = tk.Label(root, text="لم يتم اختيار ملف بعد", fg="gray", font=("Arial", 10))
        self.file_label.pack(pady=5)
        
        btn_select = tk.Button(root, text="📂 اختيار ملف الوورد (المدخلات)", command=self.select_file, width=30, font=("Arial", 11), bg="#e0e0e0")
        btn_select.pack(pady=5)
        
        # 3. زر بدء العملية
        self.btn_start = tk.Button(root, text="⚙️ بدء التنظيم والفرز", command=self.start_processing, width=30, font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", state=tk.DISABLED)
        self.btn_start.pack(pady=15)
        
        # 4. شريط حالة (عرض النتائج تحت)
        self.status_label = tk.Label(root, text="", font=("Arial", 11), fg="blue")
        self.status_label.pack(pady=5)
        
    def select_file(self):
        """نافذة لاختيار الملف المراد تنظيمه"""
        file_path = filedialog.askopenfilename(
            title="اختر ملف الوورد",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.input_file_path = file_path
            self.file_label.config(text=f"تم اختيار: {file_path.split('/')[-1]}", fg="green")
            self.btn_start.config(state=tk.NORMAL) # تفعيل زر البدء
            self.status_label.config(text="")

    def start_processing(self):
        """بدء العملية وطلب مكان حفظ الملف الناتج"""
        if not self.input_file_path:
            return
            
        # نافذة لاختيار مكان واسم حفظ الملف الناتج (التحميل)
        output_file_path = filedialog.asksaveasfilename(
            title="حفظ الملف المنظم كـ",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialfile="النتيجة_النهائية_للعوائل.docx"
        )
        
        if not output_file_path:
            return # المستخدم ألغى الحفظ
            
        try:
            self.status_label.config(text="⏳ جاري معالجة البيانات، يرجى الانتظار...", fg="orange")
            self.root.update()
            
            # تشغيل الدالة الرئيسية من ملف main.py
            main.process_and_generate(self.input_file_path, output_file_path)
            
            self.status_label.config(text="✅ تمت العملية بنجاح! تم حفظ الملف.", fg="green")
            messagebox.showinfo("نجاح", f"تم ترتيب وحفظ الملف بنجاح في:\n{output_file_path}")
            
        except Exception as e:
            self.status_label.config(text="❌ حدث خطأ أثناء المعالجة!", fg="red")
            messagebox.showerror("خطأ", f"حدث خطأ غير متوقع:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FamilySorterApp(root)
    root.mainloop()