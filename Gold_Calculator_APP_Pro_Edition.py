import tkinter as tka
from tkinter import messagebox, filedialog
import json
import os
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class GoldPriceCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("حاسبة أسعار الذهب والمجوهرات")
        self.root.geometry("1000x700") 
        
        self.colors = {
            "bg": "primary",
            "header_bg": "primary",
            "frame_bg": "secondary",
            "button": "info",
            "text": "dark",
            "highlight": "primary"
        }
        
        self.fonts = {
            "title": ("Tahoma", 14, "bold"),
            "header": ("Tahoma", 12, "bold"),
            "normal": ("Tahoma", 9),
            "small": ("Tahoma", 8)
        }
        
        self.language = "ar"
        self.translations = {
            "ar": {
                "title": "حاسبة أسعار الذهب والمجوهرات",
                "company": "حاسبة أسعار الذهب والمجوهرات للموظفين",
                "gold_price_tab": "حاسبة الذهب",
                "diamond_price_tab": "حاسبة الألماس",
                "gold_price": "سعر الذهب",
                "carat": "العيار:",
                "piece_data": "بيانات القطعة",
                "weight": "الوزن (جرام):",
                "manufacturing": "المصنعية (درهم/جرام):",
                "discount": "الخصم (%):",
                "suggested_price": "السعر المقترح (درهم):", 
                "calculate": "حساب السعر",
                "clear": "مسح الحقول",
                "save": "حفظ النتائج",
                "export_excel": "تصدير لإكسل",
                "customer_results": "نتائج العميل",
                "company_results": "نتائج الشركة",
                "piece_cost": "تكلفة القطعة:",
                "price_before_discount": "السعر قبل الخصم:",
                "discount_amount": "قيمة الخصم:",
                "price_after_discount": "السعر بعد الخصم:",
                "vat": "الضريبة (4.7619%):",
                "final_price": "السعر النهائي:",
                "net_price": "صافي السعر:",
                "net_profit": "صافي الربح:",
                "gross_margin": "هامش الربح:",
                "manufacturing_cost_label": "المصنعية لكل جرام:", 
                "language": "اللغة",
                "arabic": "العربية",
                "english": "الإنجليزية",
                "error": "خطأ",
                "success": "نجاح",
                "no_results_save": "لا توجد نتائج للحفظ",
                "no_results_export": "لا توجد نتائج للتصدير",
                "save_success": "تم حفظ النتائج في:",
                "save_fail": "فشل في الحفظ:",
                "export_success": "تم التصدير إلى:",
                "export_fail": "فشل في التصدير:",
                "positive_values_error": "القيم يجب أن تكون موجبة",
                "calculation_success": "تم الحساب بنجاح",
                "input_error": "خطأ في المدخلات",
                "ready_calculate": "جاهز للحساب",
                "no_results_yet": "لا توجد نتائج بعد",
                "theme": "السمة",
                "tag_price": "سعر الباركود (درهم):",
                "cost_tag_divided_by": "Cost (Tag divided by):",
                "calculated_cost": "التكلفة المحسوبة:",
                "diamond_customer_price": "سعر العميل (درهم):",
                "proposed_discount": "الخصم المقترح (%):",
                "item_cost": "تكلفة القطعة:",
                "diamond_tax": "الضريبة (4.7619%):",
                "diamond_final_price": "السعر النهائي (العميل):",
                "diamond_net_price": "صافي السعر (الشركة):",
                "diamond_gross_profit": "إجمالي الربح (الشركة):",
                "diamond_gross_margin": "هامش الربح الإجمالي:",
                "calc_tip": "💡 يمكنك كتابة عمليات حسابية مثل: 100*2.5 أو 50+30",
            },
            "en": {
                "title": "Jewellery Price Calculator",
                "company": "Jewellery Price Calculator for staff",
                "gold_price_tab": "Gold Calculator",
                "diamond_price_tab": "Diamond Calculator",
                "gold_price": "Gold Price",
                "carat": "Carat:",
                "piece_data": "Piece Data",
                "weight": "Weight (grams):",
                "manufacturing": "Manufacturing (AED/gram):",
                "discount": "Discount (%):",
                "suggested_price": "Suggested Price (AED):", 
                "calculate": "Calculate Price",
                "clear": "Clear Fields",
                "save": "Save Results",
                "export_excel": "Export to Excel",
                "customer_results": "Customer Results",
                "company_results": "Company Results",
                "piece_cost": "Piece Cost:",
                "price_before_discount": "Price Before Discount:",
                "discount_amount": "Discount Amount:",
                "price_after_discount": "Price After Discount:",
                "vat": "VAT (4.7619%):",
                "final_price": "Final Price:",
                "net_price": "Net Price:",
                "net_profit": "Net Profit:",
                "gross_margin": "Gross Margin:",
                "manufacturing_cost_label": "Manufacturing Cost per Gram:",
                "language": "Language",
                "arabic": "Arabic",
                "english": "English",
                "error": "Error",
                "success": "Success",
                "no_results_save": "No results to save",
                "no_results_export": "No results to export",
                "save_success": "Results saved to:",
                "save_fail": "Failed to save:",
                "export_success": "Exported to:",
                "export_fail": "Failed to export:",
                "positive_values_error": "Values must be positive",
                "calculation_success": "Calculation successful",
                "input_error": "Input error",
                "ready_calculate": "Ready to calculate",
                "no_results_yet": "No results yet",
                "theme": "Theme",
                "tag_price": "Barcode Price (AED):",
                "cost_tag_divided_by": "Cost (Tag divided by):",
                "calculated_cost": "Calculated Cost:",
                "diamond_customer_price": "Customer Price (AED):",
                "proposed_discount": "Proposed Discount (%):",
                "item_cost": "Item Cost:",
                "diamond_tax": "VAT (4.7619%):",
                "diamond_final_price": "Final Price (Customer):",
                "diamond_net_price": "Net Price (Company):",
                "diamond_gross_profit": "Gross Profit (Company):",
                "diamond_gross_margin": "Gross Margin:",
                "calc_tip": "💡 You can type calculations like: 100*2.5 or 50+30",
            }
        }
        
        # Gold Variables
        self.gold_price_var = tka.DoubleVar(value=0)
        self.weight_var = tka.DoubleVar()
        self.manufacturing_var = tka.DoubleVar()
        self.discount_var = tka.DoubleVar(value=0) 
        self.suggested_price_gold_var = tka.StringVar(value="0.0")
        self.carat_var = tka.StringVar(value="18")

        # Diamond Variables
        self.tag_price_var = tka.DoubleVar(value=0)
        self.cost_divisor_var = tka.DoubleVar(value=1.5)
        self.customer_price_diamond_var = tka.StringVar(value="") 
        self.proposed_discount_diamond_var = tka.StringVar(value="")
        
        self._updating_diamond_fields = False

        # Track current results and active tab
        self.current_gold_results = None
        self.current_diamond_results = None
        self.current_active_tab = "gold"

        self.available_themes = [
            "litera", "minty", "pulse", "flatly", "journal", 
            "lumen", "sandstone", "united", "yeti", "cosmo",
            "simplex", "cerculean", "superhero", "darkly", "solar",
            "cyborg", "vapor", "morph", "materia", "atomic",
            "azure", "forest"
        ]
        
        self.create_main_interface()
    
    def evaluate_expression(self, expression):
        """تقييم التعبيرات الرياضية البسيطة (جمع، طرح، ضرب، قسمة)"""
        if not expression or not isinstance(expression, str):
            return None
        
        expression = expression.strip()
        
        # لو الرقم فقط بدون عملية
        try:
            return float(expression)
        except ValueError:
            pass
        
        # السماح فقط بالأرقام والعمليات الأساسية
        allowed_chars = set('0123456789.-+*/')
        if not all(c in allowed_chars for c in expression):
            return None
        
        try:
            # استخدام eval آمن مع التعبيرات البسيطة
            result = eval(expression)
            return float(result)
        except:
            return None
    
    def setup_calculator_entry(self, entry_widget, target_variable):
        """إعداد خانة الإدخال لدعم العمليات الحسابية"""
        def on_focus_out(event):
            """عند مغادرة الحقل - حساب التعبير"""
            current_value = entry_widget.get().strip()
            if current_value.startswith('='):
                # لو كتب المستخدم = في البداية
                expression = current_value[1:]
                result = self.evaluate_expression(expression)
                if result is not None:
                    entry_widget.delete(0, tka.END)
                    entry_widget.insert(0, f"{result:.2f}")
                    target_variable.set(result)
                else:
                    # لو التعبير خطأ، نرجع للقيمة القديمة
                    entry_widget.delete(0, tka.END)
                    entry_widget.insert(0, str(target_variable.get()))
            else:
                # محاولة تقييم كتعبير مباشر بدون =
                result = self.evaluate_expression(current_value)
                if result is not None and current_value != str(target_variable.get()):
                    entry_widget.delete(0, tka.END)
                    entry_widget.insert(0, f"{result:.2f}")
                    target_variable.set(result)
        
        def on_key_press(event):
            """عند الضغط على Enter"""
            if event.keysym == 'Return':
                on_focus_out(event)
        
        entry_widget.bind('<FocusOut>', on_focus_out)
        entry_widget.bind('<Return>', on_key_press)
    
    def add_calculator_tooltip(self, widget):
        """إضافة تلميح للمستخدم عن إمكانية العمليات الحسابية"""
        def show_tooltip(event):
            self.tooltip_label = ttk.Label(widget, text=self.translate("calc_tip"), 
                                           bootstyle="info", font=self.fonts["small"])
            self.tooltip_label.place(x=event.x_root - widget.winfo_rootx(), 
                                     y=event.y_root - widget.winfo_rooty() + 20)
        
        def hide_tooltip(event):
            if hasattr(self, 'tooltip_label'):
                self.tooltip_label.destroy()
        
        widget.bind('<Enter>', show_tooltip)
        widget.bind('<Leave>', hide_tooltip)
    
    def translate(self, key):
        """Returns the translated string for a given key."""
        return self.translations[self.language].get(key, key)
    
    def change_language(self, lang):
        """Changes the application language and updates UI text."""
        self.language = lang
        self.update_ui_text()
        if self.current_active_tab == "gold" and self.current_gold_results: 
            self.display_gold_results()
        elif self.current_active_tab == "diamond" and self.current_diamond_results:
            self.display_diamond_results()
        else: 
            self.clear_all_results_display()

    def change_theme(self, theme_name):
        """Changes the application theme."""
        self.root.style.theme_use(theme_name) 
        self.status_var.set(f"تم تغيير السمة إلى: {theme_name}" if self.language == "ar" else f"Theme changed to: {theme_name}")

    def create_main_interface(self):
        """Creates the main application interface with tabs."""
        menubar = ttk.Menu(self.root)
        
        self.language_menu = ttk.Menu(menubar, tearoff=0)
        self.language_menu.add_command(label=self.translations["ar"]["arabic"], 
                                         command=lambda: self.change_language("ar"))
        self.language_menu.add_command(label=self.translations["en"]["english"], 
                                         command=lambda: self.change_language("en"))
        menubar.add_cascade(label=self.translate("language"), menu=self.language_menu)

        self.theme_menu = ttk.Menu(menubar, tearoff=0)
        for theme in self.available_themes:
            self.theme_menu.add_command(label=theme.capitalize(), command=lambda t=theme: self.change_theme(t))
        menubar.add_cascade(label=self.translate("theme"), menu=self.theme_menu) 
        
        self.root.config(menu=menubar)
        
        self.header = ttk.Label(self.root, text=self.translate("company"), 
                                 font=self.fonts["header"], 
                                 bootstyle=(self.colors["header_bg"], INVERSE),
                                 padding=10)
        self.header.pack(fill="x", padx=10, pady=10)
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # --- Gold Calculator Tab ---
        self.gold_tab_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.gold_tab_frame, text=self.translate("gold_price_tab"))
        self.create_gold_calculator_ui(self.gold_tab_frame)

        # --- Diamond Calculator Tab ---
        self.diamond_tab_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.diamond_tab_frame, text=self.translate("diamond_price_tab"))
        self.create_diamond_calculator_ui(self.diamond_tab_frame)
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        self.status_var = tka.StringVar()
        self.status_bar = ttk.Label(self.root, 
                                     textvariable=self.status_var, 
                                     bootstyle=(self.colors["header_bg"], INVERSE),
                                     anchor="w",
                                     relief="sunken",
                                     font=self.fonts["small"])
        self.status_bar.pack(fill="x")
        
        self.status_var.set(self.translate("ready_calculate"))

    def create_gold_calculator_ui(self, parent_frame):
        """Creates the UI elements for the Gold Calculator tab."""
        top_frames_container = ttk.Frame(parent_frame)
        top_frames_container.pack(fill="x", padx=5, pady=5)
        
        # Gold Price Frame
        self.gold_price_frame = ttk.LabelFrame(top_frames_container, text=self.translate("gold_price"), 
                                                 padding=15,
                                                 bootstyle=self.colors["frame_bg"])
        self.gold_price_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        top_frames_container.grid_columnconfigure(0, weight=1) 
        
        self.carat_label = ttk.Label(self.gold_price_frame, text=self.translate("carat"), 
                                       font=self.fonts["normal"], anchor="e")
        self.carat_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        carat_options = ttk.Combobox(self.gold_price_frame, textvariable=self.carat_var, 
                                         values=["21", "18"], 
                                         state="readonly",
                                         font=self.fonts["normal"],
                                         bootstyle="primary")
        carat_options.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        self.price_label = ttk.Label(self.gold_price_frame, 
                                         text=f"{self.translate('gold_price')} (AED/gram):", 
                                         font=self.fonts["normal"], anchor="e")
        self.price_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        
        self.price_entry = ttk.Entry(self.gold_price_frame, 
                                         textvariable=self.gold_price_var,
                                         font=self.fonts["normal"],
                                         justify="right",
                                         bootstyle="info")
        self.price_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.price_entry, self.gold_price_var)
        self.add_calculator_tooltip(self.price_entry)
        
        # Gold Piece Data Input Frame
        self.input_frame = ttk.LabelFrame(top_frames_container, text=self.translate("piece_data"), 
                                             padding=15,
                                             bootstyle=self.colors["frame_bg"])
        self.input_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5) 
        top_frames_container.grid_columnconfigure(1, weight=1) 
        
        self.weight_label = ttk.Label(self.input_frame, text=self.translate("weight"), 
                                       font=self.fonts["normal"], anchor="e")
        self.weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        
        self.weight_entry = ttk.Entry(self.input_frame, 
                  textvariable=self.weight_var,
                  font=self.fonts["normal"],
                  justify="right",
                  bootstyle="info")
        self.weight_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.weight_entry, self.weight_var)
        self.add_calculator_tooltip(self.weight_entry)
        
        self.manufacturing_label = ttk.Label(self.input_frame, 
                                             text=self.translate("manufacturing"), 
                                             font=self.fonts["normal"], anchor="e")
        self.manufacturing_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        
        self.manufacturing_entry = ttk.Entry(self.input_frame, 
                  textvariable=self.manufacturing_var,
                  font=self.fonts["normal"],
                  justify="right",
                  bootstyle="info")
        self.manufacturing_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.manufacturing_entry, self.manufacturing_var)
        self.add_calculator_tooltip(self.manufacturing_entry)
        
        self.suggested_price_label = ttk.Label(self.input_frame, text=self.translate("suggested_price"),
                                                 font=self.fonts["normal"], anchor="e")
        self.suggested_price_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        
        self.suggested_price_entry = ttk.Entry(self.input_frame, 
                                                 textvariable=self.suggested_price_gold_var,
                                                 font=self.fonts["normal"],
                                                 justify="right",
                                                 bootstyle="info")
        self.suggested_price_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.suggested_price_entry, self.suggested_price_gold_var)
        self.add_calculator_tooltip(self.suggested_price_entry)
        self.suggested_price_gold_var.trace_add("write", self.update_gold_discount_field_state)

        self.discount_label = ttk.Label(self.input_frame, text=self.translate("discount"), 
                                         font=self.fonts["normal"], anchor="e")
        self.discount_label.grid(row=3, column=0, padx=5, pady=5, sticky="e") 
        
        self.discount_entry = ttk.Entry(self.input_frame, 
                                         textvariable=self.discount_var,
                                         font=self.fonts["normal"],
                                         justify="right",
                                         bootstyle="info")
        self.discount_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.discount_entry, self.discount_var)
        self.add_calculator_tooltip(self.discount_entry)
        
        # Buttons for Gold Calculator
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(pady=15)
        
        self.calculate_gold_btn = ttk.Button(btn_frame, 
                                             text=self.translate("calculate"), 
                                             command=self.calculate_gold_price, 
                                             bootstyle="primary",
                                             cursor="hand2")
        self.calculate_gold_btn.pack(side="right", padx=10)
        
        self.clear_gold_btn = ttk.Button(btn_frame, 
                                        text=self.translate("clear"), 
                                        command=self.clear_gold_fields,
                                        bootstyle="secondary-outline",
                                        cursor="hand2")
        self.clear_gold_btn.pack(side="left", padx=10)
        
        self.save_gold_btn = ttk.Button(btn_frame, 
                                       text=self.translate("save"), 
                                       command=self.save_results,
                                       bootstyle="success",
                                       cursor="hand2")
        self.save_gold_btn.pack(side="left", padx=10)
        
        self.export_gold_btn = ttk.Button(btn_frame, 
                                          text=self.translate("export_excel"), 
                                          command=self.export_to_excel,
                                          bootstyle="info",
                                          cursor="hand2")
        self.export_gold_btn.pack(side="left", padx=10)
        
        # Results Frame
        results_container = ttk.Frame(parent_frame)
        results_container.pack(fill="both", expand=True, padx=5, pady=10)
        
        self.gold_customer_frame = ttk.LabelFrame(results_container, 
                                                    text=self.translate("customer_results"), 
                                                    padding=15,
                                                    bootstyle=self.colors["frame_bg"])
        self.gold_customer_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        self.gold_company_frame = ttk.LabelFrame(results_container, 
                                                   text=self.translate("company_results"), 
                                                   padding=15,
                                                   bootstyle=self.colors["frame_bg"])
        self.gold_company_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        
        self.clear_gold_results_display()

    def create_diamond_calculator_ui(self, parent_frame):
        """Creates the UI elements for the Diamond Calculator tab."""
        diamond_input_outer_frame = ttk.Frame(parent_frame)
        diamond_input_outer_frame.pack(fill="x", padx=5, pady=5)

        self.diamond_input_frame = ttk.LabelFrame(diamond_input_outer_frame, text=self.translate("piece_data"),
                                                  padding=15,
                                                  bootstyle=self.colors["frame_bg"])
        self.diamond_input_frame.pack(fill="x", padx=5, pady=5)
        
        # Tag Price
        self.tag_price_label = ttk.Label(self.diamond_input_frame, text=self.translate("tag_price"),
                                           font=self.fonts["normal"], anchor="e")
        self.tag_price_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.tag_price_entry = ttk.Entry(self.diamond_input_frame,
                  textvariable=self.tag_price_var,
                  font=self.fonts["normal"],
                  justify="right",
                  bootstyle="info")
        self.tag_price_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.tag_price_entry, self.tag_price_var)
        self.add_calculator_tooltip(self.tag_price_entry)

        # Cost (Tag divided by) dropdown
        self.cost_divisor_label = ttk.Label(self.diamond_input_frame, text=self.translate("cost_tag_divided_by"),
                                            font=self.fonts["normal"], anchor="e")
        self.cost_divisor_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        cost_divisor_options = ttk.Combobox(self.diamond_input_frame, textvariable=self.cost_divisor_var,
                                            values=[1.5, 1.75, 2, 2.25, 2.5, 3, 4],
                                            state="readonly",
                                            font=self.fonts["normal"],
                                            bootstyle="primary")
        cost_divisor_options.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.cost_divisor_var.set(1.5)

        # Customer Price (input/output)
        self.customer_price_diamond_label = ttk.Label(self.diamond_input_frame, text=self.translate("diamond_customer_price"),
                                                      font=self.fonts["normal"], anchor="e")
        self.customer_price_diamond_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.customer_price_diamond_entry = ttk.Entry(self.diamond_input_frame,
                                                      textvariable=self.customer_price_diamond_var,
                                                      font=self.fonts["normal"],
                                                      justify="right",
                                                      bootstyle="info")
        self.customer_price_diamond_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.customer_price_diamond_entry, self.customer_price_diamond_var)
        self.add_calculator_tooltip(self.customer_price_diamond_entry)
        self.customer_price_diamond_var.trace_add("write", self.on_diamond_customer_price_change)

        # Proposed Discount (%) (input/output)
        self.proposed_discount_label = ttk.Label(self.diamond_input_frame, text=self.translate("proposed_discount"),
                                               font=self.fonts["normal"], anchor="e")
        self.proposed_discount_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.proposed_discount_diamond_entry = ttk.Entry(self.diamond_input_frame,
                                                        textvariable=self.proposed_discount_diamond_var,
                                                        font=self.fonts["normal"],
                                                        justify="right",
                                                        bootstyle="info")
        self.proposed_discount_diamond_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.setup_calculator_entry(self.proposed_discount_diamond_entry, self.proposed_discount_diamond_var)
        self.add_calculator_tooltip(self.proposed_discount_diamond_entry)
        self.proposed_discount_diamond_var.trace_add("write", self.on_diamond_proposed_discount_change)

        # Buttons for Diamond Calculator
        btn_frame = ttk.Frame(parent_frame)
        btn_frame.pack(pady=15)
        
        self.calculate_diamond_btn = ttk.Button(btn_frame, 
                                                text=self.translate("calculate"), 
                                                command=self.calculate_diamond_price, 
                                                bootstyle="primary",
                                                cursor="hand2")
        self.calculate_diamond_btn.pack(side="right", padx=10)
        
        self.clear_diamond_btn = ttk.Button(btn_frame, 
                                            text=self.translate("clear"), 
                                            command=self.clear_diamond_fields,
                                            bootstyle="secondary-outline",
                                            cursor="hand2")
        self.clear_diamond_btn.pack(side="left", padx=10)
        
        self.save_diamond_btn = ttk.Button(btn_frame, 
                                           text=self.translate("save"), 
                                           command=self.save_results,
                                           bootstyle="success",
                                           cursor="hand2")
        self.save_diamond_btn.pack(side="left", padx=10)
        
        self.export_diamond_btn = ttk.Button(btn_frame, 
                                             text=self.translate("export_excel"), 
                                             command=self.export_to_excel,
                                             bootstyle="info",
                                             cursor="hand2")
        self.export_diamond_btn.pack(side="left", padx=10)

        # Results Frame for Diamond
        diamond_results_container = ttk.Frame(parent_frame)
        diamond_results_container.pack(fill="both", expand=True, padx=5, pady=10)

        self.diamond_customer_frame = ttk.LabelFrame(diamond_results_container,
                                                      text=self.translate("customer_results"),
                                                      padding=15,
                                                      bootstyle=self.colors["frame_bg"])
        self.diamond_customer_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        self.diamond_company_frame = ttk.LabelFrame(diamond_results_container,
                                                     text=self.translate("company_results"),
                                                     padding=15,
                                                     bootstyle=self.colors["frame_bg"])
        self.diamond_company_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        self.clear_diamond_results_display()

    def on_tab_change(self, event):
        """Handles actions when the notebook tab is changed."""
        selected_tab_id = self.notebook.tab(self.notebook.select(), "text")
        
        if selected_tab_id == self.translate("gold_price_tab"):
            self.current_active_tab = "gold"
        elif selected_tab_id == self.translate("diamond_price_tab"):
            self.current_active_tab = "diamond"
        
        self.clear_all_results_display()
        self.status_var.set(self.translate("ready_calculate"))

    def update_gold_discount_field_state(self, *args):
        """Manages the state of the discount entry for gold calculation based on suggested price."""
        suggested_price_str = self.suggested_price_gold_var.get().strip() 
        
        try:
            if suggested_price_str and float(suggested_price_str) > 0: 
                self.discount_entry.config(state="readonly")
                self.discount_var.set(0.0)
            else:
                self.discount_entry.config(state="normal")
        except ValueError:
            self.discount_entry.config(state="normal")
            self.discount_var.set(0.0) 
            
    def on_diamond_customer_price_change(self, *args):
        """Calculates proposed discount when customer price changes and manages field states."""
        if self._updating_diamond_fields:
            return

        self._updating_diamond_fields = True
        
        tag_price_str = self.tag_price_var.get()
        customer_price_str = self.customer_price_diamond_var.get().strip()

        try:
            tag_price = float(tag_price_str)
            customer_price = float(customer_price_str)
            
            if tag_price > 0 and customer_price >= 0:
                calculated_discount = ((tag_price - customer_price) / tag_price) * 100
                self.proposed_discount_diamond_var.set(f"{max(0, calculated_discount):.2f}")
                self.proposed_discount_diamond_entry.config(state="readonly")
                self.customer_price_diamond_entry.config(state="normal")
            else:
                self.proposed_discount_diamond_var.set("")
                self.proposed_discount_diamond_entry.config(state="normal")
        except ValueError:
            self.proposed_discount_diamond_var.set("")
            self.proposed_discount_diamond_entry.config(state="normal")
            self.customer_price_diamond_entry.config(state="normal")

        self._updating_diamond_fields = False

    def on_diamond_proposed_discount_change(self, *args):
        """Calculates customer price when proposed discount changes and manages field states."""
        if self._updating_diamond_fields:
            return

        self._updating_diamond_fields = True

        tag_price_str = self.tag_price_var.get()
        proposed_discount_str = self.proposed_discount_diamond_var.get().strip()

        try:
            tag_price = float(tag_price_str)
            proposed_discount = float(proposed_discount_str)

            if tag_price > 0 and 0 <= proposed_discount <= 100:
                calculated_customer_price = tag_price * (1 - (proposed_discount / 100))
                self.customer_price_diamond_var.set(f"{max(0, calculated_customer_price):.2f}")
                self.customer_price_diamond_entry.config(state="readonly")
                self.proposed_discount_diamond_entry.config(state="normal")
            else:
                self.customer_price_diamond_var.set("")
                self.customer_price_diamond_entry.config(state="normal")
        except ValueError:
            self.customer_price_diamond_var.set("")
            self.customer_price_diamond_entry.config(state="normal")
            self.proposed_discount_diamond_entry.config(state="normal")

        self._updating_diamond_fields = False

    def update_ui_text(self):
        """Updates all UI text elements based on the current language."""
        self.root.title(self.translate("title"))
        self.header.config(text=self.translate("company"))
        
        self.notebook.tab(0, text=self.translate("gold_price_tab"))
        self.notebook.tab(1, text=self.translate("diamond_price_tab"))

        # Gold UI updates
        self.gold_price_frame.config(text=self.translate("gold_price"))
        self.carat_label.config(text=self.translate("carat"))
        self.price_label.config(text=f"{self.translate('gold_price')} (AED/gram):")
        self.input_frame.config(text=self.translate("piece_data"))
        self.weight_label.config(text=self.translate("weight"))
        self.manufacturing_label.config(text=self.translate("manufacturing"))
        self.discount_label.config(text=self.translate("discount"))
        self.suggested_price_label.config(text=self.translate("suggested_price"))
        self.calculate_gold_btn.config(text=self.translate("calculate"))
        self.clear_gold_btn.config(text=self.translate("clear"))
        self.save_gold_btn.config(text=self.translate("save"))
        self.export_gold_btn.config(text=self.translate("export_excel"))
        self.gold_customer_frame.config(text=self.translate("customer_results"))
        self.gold_company_frame.config(text=self.translate("company_results"))
        
        # Diamond UI updates
        self.diamond_input_frame.config(text=self.translate("piece_data"))
        self.tag_price_label.config(text=self.translate("tag_price"))
        self.cost_divisor_label.config(text=self.translate("cost_tag_divided_by"))
        self.customer_price_diamond_label.config(text=self.translate("diamond_customer_price"))
        self.proposed_discount_label.config(text=self.translate("proposed_discount"))

        self.calculate_diamond_btn.config(text=self.translate("calculate"))
        self.clear_diamond_btn.config(text=self.translate("clear"))
        self.save_diamond_btn.config(text=self.translate("save"))
        self.export_diamond_btn.config(text=self.translate("export_excel"))
        self.diamond_customer_frame.config(text=self.translate("customer_results"))
        self.diamond_company_frame.config(text=self.translate("company_results"))
        
        self.status_var.set(self.translate("ready_calculate"))
        
        if self.current_active_tab == "gold" and self.current_gold_results:
            self.display_gold_results()
        elif self.current_active_tab == "diamond" and self.current_diamond_results:
            self.display_diamond_results()
        else:
            self.clear_all_results_display() 
    
    def calculate_gold_price(self):
        """Calculates prices for gold based on inputs."""
        try:
            gold_price = self.gold_price_var.get()
            weight = self.weight_var.get()
            manufacturing = self.manufacturing_var.get()
            carat = self.carat_var.get()
            
            suggested_price_str = self.suggested_price_gold_var.get().strip()
            if suggested_price_str:
                suggested_price = float(suggested_price_str)
            else:
                suggested_price = 0.0

            if gold_price <= 0 or weight <= 0 or manufacturing < 0:
                raise ValueError(self.translate("positive_values_error"))
            
            conversion_factor = float(carat) / 24
            adjusted_gold_price = gold_price * 1
            
            piece_cost = ((manufacturing / 3) + adjusted_gold_price) * weight
            
            if manufacturing > 60:
                customer_price_before_discount = ((manufacturing + adjusted_gold_price) * weight) * 1.3
            else:
                customer_price_before_discount = ((manufacturing + adjusted_gold_price) * weight) * 1.5
            
            discount_value_for_calculation = 0 
            
            if suggested_price > 0:
                customer_price_after_discount_calculated = suggested_price
                if customer_price_before_discount > 0:
                    calculated_discount_amount = customer_price_before_discount - suggested_price
                    if calculated_discount_amount < 0:
                        calculated_discount_amount = 0
                    calculated_discount_percentage = (calculated_discount_amount / customer_price_before_discount) * 100
                else:
                    calculated_discount_amount = 0
                    calculated_discount_percentage = 0

                self.discount_var.set(round(calculated_discount_percentage, 2)) 
                discount_value_for_calculation = calculated_discount_percentage
            else:
                discount_value_for_calculation = self.discount_var.get()
                if discount_value_for_calculation < 0:
                    raise ValueError(self.translate("positive_values_error"))
                
                discount_amount = customer_price_before_discount * (discount_value_for_calculation / 100)
                customer_price_after_discount_calculated = customer_price_before_discount - discount_amount
            
            discount_amount = customer_price_before_discount - customer_price_after_discount_calculated

            vat_rate = 0.047619
            vat_amount = customer_price_after_discount_calculated * vat_rate
            
            customer_price_with_vat = customer_price_after_discount_calculated
            
            net_price_for_company = customer_price_after_discount_calculated - vat_amount

            net_profit = net_price_for_company - piece_cost
            
            if net_price_for_company > 0:
                gross_margin = (net_profit / net_price_for_company) * 100
            else:
                gross_margin = 0

            if weight > 0 and gold_price > 0:
                manufacturing_cost_per_gram = (net_price_for_company / weight) - adjusted_gold_price
            else:
                manufacturing_cost_per_gram = 0 
            
            self.current_gold_results = {
                "type": "gold",
                "gold_price": gold_price,
                "adjusted_gold_price": adjusted_gold_price,
                "weight": weight,
                "manufacturing": manufacturing,
                "discount": discount_value_for_calculation,
                "suggested_price": suggested_price,
                "carat": carat,
                "piece_cost": piece_cost,
                "customer_price_before_discount": customer_price_before_discount,
                "discount_amount": discount_amount,
                "customer_price_after_discount": customer_price_after_discount_calculated,
                "vat_amount": vat_amount,
                "customer_price_with_vat": customer_price_with_vat, 
                "net_price_for_company": net_price_for_company,
                "net_profit": net_profit,
                "gross_margin": gross_margin,
                "manufacturing_cost_per_gram": manufacturing_cost_per_gram,
                "calculation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.current_active_tab = "gold"
            
            self.display_gold_results()
            
            self.status_var.set(self.translate("calculation_success"))
            
        except ValueError as ve:
            messagebox.showerror(self.translate("error"), str(ve))
            self.status_var.set(self.translate("input_error"))
        except Exception as e:
            messagebox.showerror(self.translate("error"), f"{self.translate('error')}: {str(e)}")
            self.status_var.set(self.translate("input_error"))
    
    def calculate_diamond_price(self):
        """Calculates prices for diamonds based on inputs."""
        try:
            tag_price = self.tag_price_var.get()
            cost_divisor = self.cost_divisor_var.get()
            
            customer_price_str = self.customer_price_diamond_var.get().strip()
            proposed_discount_str = self.proposed_discount_diamond_var.get().strip()

            if tag_price <= 0 or cost_divisor <= 0:
                raise ValueError(self.translate("positive_values_error"))
            
            customer_price = 0.0
            proposed_discount = 0.0

            if customer_price_str:
                customer_price = float(customer_price_str)
                if tag_price > 0:
                    proposed_discount = ((tag_price - customer_price) / tag_price) * 100
                else:
                    proposed_discount = 0
            elif proposed_discount_str:
                proposed_discount = float(proposed_discount_str)
                if 0 <= proposed_discount <= 100:
                    customer_price = tag_price * (1 - (proposed_discount / 100))
                else:
                    raise ValueError("Discount must be between 0 and 100.")
            else:
                raise ValueError("Please enter either Customer Price or Proposed Discount.")

            item_cost = tag_price / cost_divisor

            vat_rate = 0.047619
            tax_amount = customer_price * vat_rate

            final_customer_price_display = customer_price

            net_price_company = customer_price - tax_amount
            gross_profit = net_price_company - item_cost
            
            gross_margin = 0
            if net_price_company > 0:
                gross_margin = (gross_profit / net_price_company) * 100

            self.current_diamond_results = {
                "type": "diamond",
                "tag_price": tag_price,
                "cost_divisor": cost_divisor,
                "customer_price": customer_price, 
                "proposed_discount": proposed_discount,
                "calculated_item_cost": item_cost,
                "diamond_tax_amount": tax_amount,
                "customer_final_price_display": final_customer_price_display,
                "net_price_company": net_price_company,
                "gross_profit": gross_profit,
                "gross_margin": gross_margin,
                "calculation_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.current_active_tab = "diamond"

            self.display_diamond_results()
            self.status_var.set(self.translate("calculation_success"))

        except ValueError as ve:
            messagebox.showerror(self.translate("error"), str(ve))
            self.status_var.set(self.translate("input_error"))
        except Exception as e:
            messagebox.showerror(self.translate("error"), f"{self.translate('error')}: {str(e)}")
            self.status_var.set(self.translate("input_error"))

    def display_gold_results(self):
        """Displays the calculated gold results in the respective frames."""
        for widget in self.gold_customer_frame.winfo_children():
            widget.destroy()
        for widget in self.gold_company_frame.winfo_children():
            widget.destroy()
        
        if not self.current_gold_results:
            self.clear_gold_results_display()
            return
            
        formatted_results = {k: round(v, 2) if isinstance(v, (int, float)) else v 
                             for k, v in self.current_gold_results.items()}
        
        # Customer Results (Gold)
        row = 0
        ttk.Label(self.gold_customer_frame, text=self.translate("piece_cost"), 
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_customer_frame, text=f"{formatted_results['piece_cost']:.2f} AED",
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_customer_frame, text=self.translate("price_before_discount"), 
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_customer_frame, text=f"{formatted_results['customer_price_before_discount']:.2f} AED",
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_customer_frame, text=self.translate("discount_amount"), 
                  font=self.fonts["small"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_customer_frame, text=f"{formatted_results['discount_amount']:.2f} AED",
                  font=self.fonts["small"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_customer_frame, text=self.translate("price_after_discount"), 
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_customer_frame, text=f"{formatted_results['customer_price_after_discount']:.2f} AED",
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_customer_frame, text=self.translate("vat"), 
                  font=self.fonts["small"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_customer_frame, text=f"{formatted_results['vat_amount']:.2f} AED",
                  font=self.fonts["small"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_customer_frame, text=self.translate("final_price"), 
                  font=self.fonts["header"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_customer_frame, text=f"{formatted_results['customer_price_with_vat']:.2f} AED", 
                  bootstyle="primary", 
                  font=self.fonts["header"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        
        # Company Results (Gold)
        row = 0
        ttk.Label(self.gold_company_frame, text=self.translate("net_price"), 
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_company_frame, text=f"{formatted_results['net_price_for_company']:.2f} AED",
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_company_frame, text=self.translate("net_profit"), 
                  font=self.fonts["header"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_company_frame, text=f"{formatted_results['net_profit']:.2f} AED", 
                  bootstyle="success" if formatted_results['net_profit'] >= 0 else "danger",
                  font=self.fonts["header"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1
        
        ttk.Label(self.gold_company_frame, text=self.translate("gross_margin"), 
                  font=self.fonts["header"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_company_frame, text=f"{formatted_results['gross_margin']:.2f}%", 
                  bootstyle="success" if formatted_results['gross_margin'] >= 0 else "danger",
                  font=self.fonts["header"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)

        row += 1
        ttk.Label(self.gold_company_frame, text=self.translate("manufacturing_cost_label"), 
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.gold_company_frame, text=f"{formatted_results['manufacturing_cost_per_gram']:.2f} AED/gram", 
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
    
    def display_diamond_results(self):
        """Displays the calculated diamond results in the respective frames."""
        for widget in self.diamond_customer_frame.winfo_children():
            widget.destroy()
        for widget in self.diamond_company_frame.winfo_children():
            widget.destroy()
        
        if not self.current_diamond_results:
            self.clear_diamond_results_display()
            return
        
        formatted_results = {k: round(v, 2) if isinstance(v, (int, float)) else v 
                             for k, v in self.current_diamond_results.items()}

        # Customer Results (Diamond)
        row = 0
        ttk.Label(self.diamond_customer_frame, text=self.translate("item_cost"),
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.diamond_customer_frame, text=f"{formatted_results['calculated_item_cost']:.2f} AED",
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1

        ttk.Label(self.diamond_customer_frame, text=self.translate("diamond_tax"),
                  font=self.fonts["small"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.diamond_customer_frame, text=f"{formatted_results['diamond_tax_amount']:.2f} AED",
                  font=self.fonts["small"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1

        ttk.Label(self.diamond_customer_frame, text=self.translate("diamond_final_price"),
                  font=self.fonts["header"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.diamond_customer_frame, text=f"{formatted_results['customer_final_price_display']:.2f} AED",
                  bootstyle="primary",
                  font=self.fonts["header"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)

        # Company Results (Diamond)
        row = 0
        ttk.Label(self.diamond_company_frame, text=self.translate("diamond_net_price"),
                  font=self.fonts["normal"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.diamond_company_frame, text=f"{formatted_results['net_price_company']:.2f} AED",
                  font=self.fonts["normal"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1

        ttk.Label(self.diamond_company_frame, text=self.translate("diamond_gross_profit"),
                  font=self.fonts["header"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.diamond_company_frame, text=f"{formatted_results['gross_profit']:.2f} AED",
                  bootstyle="success" if formatted_results['gross_profit'] >= 0 else "danger",
                  font=self.fonts["header"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
        row += 1

        ttk.Label(self.diamond_company_frame, text=self.translate("diamond_gross_margin"),
                  font=self.fonts["header"], anchor="e").grid(row=row, column=0, sticky="e", padx=5, pady=2)
        ttk.Label(self.diamond_company_frame, text=f"{formatted_results['gross_margin']:.2f}%",
                  bootstyle="success" if formatted_results['gross_margin'] >= 0 else "danger",
                  font=self.fonts["header"], anchor="w").grid(row=row, column=1, sticky="w", padx=5, pady=2)
    
    def clear_gold_results_display(self):
        """Clears the gold results display frames."""
        for widget in self.gold_customer_frame.winfo_children():
            widget.destroy()
        for widget in self.gold_company_frame.winfo_children():
            widget.destroy()
        
        ttk.Label(self.gold_customer_frame, 
                  text=self.translate("no_results_yet"), 
                  font=self.fonts["normal"]).pack(expand=True, padx=10, pady=10) 
        
        ttk.Label(self.gold_company_frame, 
                  text=self.translate("no_results_yet"), 
                  font=self.fonts["normal"]).pack(expand=True, padx=10, pady=10) 
    
    def clear_diamond_results_display(self):
        """Clears the diamond results display frames."""
        for widget in self.diamond_customer_frame.winfo_children():
            widget.destroy()
        for widget in self.diamond_company_frame.winfo_children():
            widget.destroy()
        
        ttk.Label(self.diamond_customer_frame, 
                  text=self.translate("no_results_yet"), 
                  font=self.fonts["normal"]).pack(expand=True, padx=10, pady=10) 
        
        ttk.Label(self.diamond_company_frame, 
                  text=self.translate("no_results_yet"), 
                  font=self.fonts["normal"]).pack(expand=True, padx=10, pady=10) 

    def clear_all_results_display(self):
        """Clears results display for both tabs."""
        self.clear_gold_results_display()
        self.clear_diamond_results_display()

    def clear_gold_fields(self):
        """Clears all input fields for the Gold Calculator."""
        self.gold_price_var.set(0) 
        self.weight_var.set(0)
        self.manufacturing_var.set(0)
        self.discount_var.set(0)
        self.suggested_price_gold_var.set("") 
        self.carat_var.set("18") 
        self.update_gold_discount_field_state() 
        self.clear_gold_results_display()
        self.status_var.set(self.translate("ready_calculate"))
    
    def clear_diamond_fields(self):
        """Clears all input fields for the Diamond Calculator."""
        self.tag_price_var.set(0)
        self.cost_divisor_var.set(1.5)
        self._updating_diamond_fields = True
        self.customer_price_diamond_var.set("")
        self.proposed_discount_diamond_var.set("")
        self._updating_diamond_fields = False
        
        self.customer_price_diamond_entry.config(state="normal")
        self.proposed_discount_diamond_entry.config(state="normal")

        self.clear_diamond_results_display()
        self.status_var.set(self.translate("ready_calculate"))
    
    def save_results(self):
        """Saves the results of the currently active tab."""
        results_to_save = None
        if self.current_active_tab == "gold":
            results_to_save = self.current_gold_results
        elif self.current_active_tab == "diamond":
            results_to_save = self.current_diamond_results

        if not results_to_save:
            messagebox.showwarning(self.translate("error"), 
                                   self.translate("no_results_save"))
            return
            
        try:
            os.makedirs("saved_calculations", exist_ok=True)
            
            default_filename = f"{self.current_active_tab}_calc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            
            file_path = filedialog.asksaveasfilename(
                initialdir="saved_calculations",
                initialfile=default_filename,
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(results_to_save, f, ensure_ascii=False, indent=4)
                
                messagebox.showinfo(self.translate("success"), 
                                     f"{self.translate('save_success')}\n{file_path}")
                
        except Exception as e:
            messagebox.showerror(self.translate("error"), 
                                 f"{self.translate('save_fail')}\n{str(e)}")
    
    def export_to_excel(self):
        """Exports the results of the currently active tab to an Excel file."""
        results_to_export = None
        if self.current_active_tab == "gold":
            results_to_export = self.current_gold_results
        elif self.current_active_tab == "diamond":
            results_to_export = self.current_diamond_results

        if not results_to_export:
            messagebox.showwarning(self.translate("error"), 
                                   self.translate("no_results_export"))
            return
            
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"{self.current_active_tab.capitalize()} Calculation Results"
            
            if results_to_export["type"] == "gold":
                headers = [
                    "Parameter", "Value", "Unit"
                ]
                ws.append(headers)
                
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                
                data = [
                    ["Gold Price", results_to_export['gold_price'], "AED/gram"],
                    ["Adjusted Gold Price", results_to_export['adjusted_gold_price'], "AED/gram"],
                    ["Weight", results_to_export['weight'], "grams"],
                    ["Manufacturing", results_to_export['manufacturing'], "AED/gram"],
                    ["Discount", results_to_export['discount'], "%"],
                    ["Suggested Price", results_to_export['suggested_price'], "AED"], 
                    ["Carat", results_to_export['carat'], "K"],
                    ["Piece Cost", results_to_export['piece_cost'], "AED"],
                    ["Price Before Discount", results_to_export['customer_price_before_discount'], "AED"],
                    ["Discount Amount", results_to_export['discount_amount'], "AED"],
                    ["Price After Discount", results_to_export['customer_price_after_discount'], "AED"],
                    ["VAT Amount", results_to_export['vat_amount'], "AED"],
                    ["Final Price", results_to_export['customer_price_with_vat'], "AED"],
                    ["Net Price for Company", results_to_export['net_price_for_company'], "AED"],
                    ["Net Profit", results_to_export['net_profit'], "AED"],
                    ["Gross Margin", f"{results_to_export['gross_margin']:.2f}", "%"],
                    ["Manufacturing Cost per Gram", results_to_export['manufacturing_cost_per_gram'], "AED/gram"],
                    ["Calculation Time", results_to_export['calculation_time'], ""]
                ]
            elif results_to_export["type"] == "diamond":
                headers = [
                    "Parameter", "Value", "Unit"
                ]
                ws.append(headers)

                for cell in ws[1]:
                    cell.font = Font(bold=True)
                
                data = [
                    ["Barcode Price", results_to_export['tag_price'], "AED"],
                    ["Cost (Tag divided by)", results_to_export['cost_divisor'], ""],
                    ["Calculated Item Cost", results_to_export['calculated_item_cost'], "AED"],
                    ["Customer Price", results_to_export['customer_price'], "AED"],
                    ["Proposed Discount", f"{results_to_export['proposed_discount']:.2f}", "%"],
                    ["VAT Amount", results_to_export['diamond_tax_amount'], "AED"],
                    ["Customer Final Price", results_to_export['customer_final_price_display'], "AED"],
                    ["Net Price (Company)", results_to_export['net_price_company'], "AED"],
                    ["Gross Profit", results_to_export['gross_profit'], "AED"],
                    ["Gross Margin", f"{results_to_export['gross_margin']:.2f}", "%"],
                    ["Calculation Time", results_to_export['calculation_time'], ""]
                ]
            
            for row in data:
                ws.append(row)
            
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column].width = adjusted_width
            
            default_filename = f"{self.current_active_tab}_calc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = filedialog.asksaveasfilename(
                initialdir="saved_calculations",
                initialfile=default_filename,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if file_path:
                wb.save(file_path)
                messagebox.showinfo(self.translate("success"), 
                                     f"{self.translate('export_success')}\n{file_path}")
                
        except Exception as e:
            messagebox.showerror(self.translate("error"), 
                                 f"{self.translate('export_fail')}\n{str(e)}")

if __name__ == "__main__":
    root = ttk.Window(themename="simplex")
    app = GoldPriceCalculator(root)
    root.mainloop()
