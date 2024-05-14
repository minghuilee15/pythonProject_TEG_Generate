import tkinter as tk
from PIL import Image, ImageTk
import subprocess

class GUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("TEG Table Generate")
        self.geometry("600x400")
        self.center_window()

        self.current_frame = None

        # 加载并调整返回图标的大小
        self.back_icon = ImageTk.PhotoImage(Image.open("res/返回.png").resize((30, 30), Image.BICUBIC))

        self.show_main_menu()

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        x = (screen_width - 600) // 2
        y = (screen_height - 400) // 2

        self.geometry(f"600x400+{x}+{y}")

    def show_main_menu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        wlr_button = tk.Button(self.current_frame, text="WLR", command=self.show_wlr_submenu)
        wlr_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        pcm_button = tk.Button(self.current_frame, text="PCM", command=self.show_pcm_submenu)
        pcm_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        dr_button = tk.Button(self.current_frame, text="DR", command=self.show_dr_submenu)
        dr_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        spice_button = tk.Button(self.current_frame, text="SPICE", command=self.show_spice_submenu)
        spice_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10


    def show_wlr_submenu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        back_button = tk.Button(self.current_frame, image=self.back_icon, command=self.show_main_menu)
        back_button.pack()

        pid_button = tk.Button(self.current_frame, text="PID", command=lambda: self.run_script("pid.py"))
        pid_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        nbti_button = tk.Button(self.current_frame, text="NBTI", command=lambda: self.run_script("nbti.py"))
        nbti_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        goi_button = tk.Button(self.current_frame, text="GOI", command=lambda: self.run_script("goi.py"))
        goi_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        tddb_button = tk.Button(self.current_frame, text="TDDB", command=lambda: self.run_script("tddb.py"))
        tddb_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        hci_button = tk.Button(self.current_frame, text="HCI", command=lambda: self.run_script("hci.py"))
        hci_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        em_button = tk.Button(self.current_frame, text="EM", command=lambda: self.run_script("em.py"))
        em_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        ild_button = tk.Button(self.current_frame, text="ILD", command=lambda: self.run_script("ild.py"))
        ild_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

    def show_pcm_submenu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        back_button = tk.Button(self.current_frame, image=self.back_icon, command=self.show_main_menu)
        back_button.pack()

        pas_button = tk.Button(self.current_frame, text="Active Device", command=self.show_active_device_submenu)
        pas_button.pack(side="left", padx=10, pady=(120, 0))

        pas_button = tk.Button(self.current_frame, text="Passive Device", command=self.show_passive_device_submenu)
        pas_button.pack(side="left", padx=10, pady=(120, 0))

    def show_active_device_submenu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        back_button = tk.Button(self.current_frame, image=self.back_icon, command=self.show_main_menu)
        back_button.pack()

        mos_button = tk.Button(self.current_frame, text="MOS", command=lambda: self.run_script("mos.py"))
        mos_button.pack(side="left", padx=10, pady=(120, 0))

        field_transistor_button = tk.Button(self.current_frame, text="Field Transistor",
                                            command=lambda: self.run_script("field_transistor.py"))
        field_transistor_button.pack(side="left", padx=10, pady=(120, 0))

        diode_button = tk.Button(self.current_frame, text="Diode", command=lambda: self.run_script("diode.py"))
        diode_button.pack(side="left", padx=10, pady=(120, 0))

        bjt_button = tk.Button(self.current_frame, text="BJT", command=lambda: self.run_script("bjt.py"))
        bjt_button.pack(side="left", padx=10, pady=(120, 0))

    def show_passive_device_submenu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        back_button = tk.Button(self.current_frame, image=self.back_icon, command=self.show_main_menu)
        back_button.pack()

        res_button = tk.Button(self.current_frame, text="Res", command=lambda: self.run_script("res.py"))
        res_button.pack(side="left", padx=10, pady=(120, 0))

        cap_button = tk.Button(self.current_frame, text="Cap", command=lambda: self.run_script("cap.py"))
        cap_button.pack(side="left", padx=10, pady=(120, 0))

    def show_dr_submenu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        back_button = tk.Button(self.current_frame, image=self.back_icon, command=self.show_main_menu)
        back_button.pack()

        ct_button = tk.Button(self.current_frame, text="CT", command=lambda: self.run_script("PEX_v01.py"))
        ct_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        via_button = tk.Button(self.current_frame, text="VIA", command=lambda: self.run_script("VIA.py"))
        via_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        isolation_button = tk.Button(self.current_frame, text="Isolation", command=lambda: self.run_script("isolation.py"))
        isolation_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        Bridge_button = tk.Button(self.current_frame, text="Bridge", command=lambda: self.run_script("Bridge.py"))
        Bridge_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        Continuity_button = tk.Button(self.current_frame, text="Continuity", command=lambda: self.run_script("Continuity.py"))
        Continuity_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

    def show_spice_submenu(self):
        if self.current_frame:
            self.current_frame.destroy()

        self.current_frame = tk.Frame(self)
        self.current_frame.pack(padx=20, pady=20)

        back_button = tk.Button(self.current_frame, image=self.back_icon, command=self.show_main_menu)
        back_button.pack()

        pex_button = tk.Button(self.current_frame, text="PEX", command=lambda: self.run_script("pex.py"))
        pex_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

        lu_button = tk.Button(self.current_frame, text="Latch Up", command=lambda: self.run_script("lu.py"))
        lu_button.pack(side="left", padx=10, pady=(120, 0))  # 设置按钮距离上边界的垂直间距为10，水平间距为10

    def run_script(self, script_name):
        subprocess.Popen(["python", script_name])

if __name__ == "__main__":
    app = GUI()
    app.mainloop()