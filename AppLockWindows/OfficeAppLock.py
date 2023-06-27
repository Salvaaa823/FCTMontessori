import tkinter as tk
import win32com.client
import threading

class OfficeAppLock:
    def __init__(self, app_names):
        self.app_names = app_names
        self.selected_app = None
        self.lock_time = 0
        self.admin_password = "1234"
        self.locked_app = None
        self.lock_thread = None
        
        self.root = tk.Tk()
        self.root.title("Office App Locker")
        
        self.app_label = tk.Label(self.root, text="Selecciona una aplicación:")
        self.app_label.pack(pady=5)
        
        self.app_listbox = tk.Listbox(self.root)
        self.app_listbox.pack(padx=10, pady=5)
        self.populate_app_list()
        
        self.time_label = tk.Label(self.root, text="Tiempo de bloqueo (segundos):")
        self.time_label.pack(pady=5)
        
        self.time_entry = tk.Entry(self.root)
        self.time_entry.pack()
        
        self.start_button = tk.Button(self.root, text="Bloquear", command=self.start_lock)
        self.start_button.pack(pady=5)
        
        self.unlock_label = tk.Label(self.root, text="")
        self.unlock_label.pack(pady=5)
        
        self.password_label = tk.Label(self.root, text="Contraseña:")
        self.password_label.pack(pady=5)
        
        self.password_entry = tk.Entry(self.root, show="*")
        self.password_entry.pack()
        
        self.countdown_label = tk.Label(self.root, text="")
        self.countdown_label.pack(pady=10)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def populate_app_list(self):
        for app_name in self.app_names:
            self.app_listbox.insert(tk.END, app_name)
    
    def start_lock(self):
        selected_index = self.app_listbox.curselection()
        if selected_index:
            self.selected_app = self.app_names[selected_index[0]]
            self.lock_time = int(self.time_entry.get())
            
            self.lock_thread = threading.Thread(target=self.lock_app)
            self.lock_thread.start()
        else:
            self.unlock_label.config(text="Selecciona una aplicación.")
        
    def lock_app(self):
        self.app_listbox.config(state=tk.DISABLED)
        self.time_entry.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        
        try:
            office_app = win32com.client.Dispatch(self.selected_app)
        except:
            self.unlock_label.config(text="No se pudo abrir la aplicación.")
            self.app_listbox.config(state=tk.NORMAL)
            self.time_entry.config(state=tk.NORMAL)
            self.start_button.config(state=tk.NORMAL)
            return
        
        office_app.Visible = True
        office_app.WindowState = 3  # Maximize window
        
        remaining_time = self.lock_time
        
        while remaining_time >= 0:
            self.countdown_label.config(text=f"Tiempo restante: {remaining_time} segundos")
            remaining_time -= 1
            self.lock_thread.join(1)  # Delay de 1 segundo
            
            if remaining_time < 0:
                self.unlock_app()
                break
        
    def unlock_app(self):
        self.locked_app = None
        
        self.app_listbox.config(state=tk.NORMAL)
        self.time_entry.config(state=tk.NORMAL)
        self.start_button.config(state=tk.NORMAL)
        self.unlock_label.config(text="")
        self.countdown_label.config(text="")
        
    def on_close(self):
        self.unlock_app()
        self.root.destroy()

if __name__ == "__main__":
    app_names = ["Word.Application", "Excel.Application", "PowerPoint.Application"]
    office_app_lock = OfficeAppLock(app_names)
    office_app_lock.root.mainloop()
