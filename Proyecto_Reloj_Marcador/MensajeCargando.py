import tkinter as tk

class MensajeCargando(tk.Toplevel):
    def __init__(self):
        super().__init__()

        ancho = 300
        alto = 150
        screen_ancho = self.winfo_screenwidth()
        screen_alto = self.winfo_screenheight()
        x = (screen_ancho // 2) - (ancho // 2)
        y = (screen_alto // 2) - (alto // 2)

        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.title("Cargando")
        self.configure(bg="#ffffff")
        self.overrideredirect(True)
        label = tk.Label(
            self,
            text="Cargando. . .",
            fg="black",
            font=("Arial", 18, "bold"), 
            bg="white"
        )
        label.pack(expand=True)