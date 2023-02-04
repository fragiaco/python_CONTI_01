import tkinter as tk


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        # Imposta il titolo della finestra principale
        self.title("Multi-Frame Tkinter Example")

        # Crea il frame di benvenuto
        self.welcome_frame = tk.Frame(self)
        self.welcome_frame.pack()

        # Crea una label di benvenuto
        tk.Label(self.welcome_frame, text="Benvenuti nella finestra principale!").pack()

        # Crea un pulsante per andare al prossimo frame
        tk.Button(self.welcome_frame, text="Vai al prossimo frame", command=self.show_next_frame).pack()

        # Crea il frame successivo
        self.next_frame = tk.Frame(self)

        # Crea una label di avviso
        tk.Label(self.next_frame, text="Questo è il prossimo frame.").pack()

        # Crea un pulsante per tornare indietro
        tk.Button(self.next_frame, text="Torna indietro", command=self.show_welcome_frame).pack()

    def show_next_frame(self):
        # Nasconde il frame di benvenuto e mostra il prossimo frame
        self.welcome_frame.pack_forget()
        self.next_frame.pack()

    def show_welcome_frame(self):
        # Nasconde il prossimo frame e mostra il frame di benvenuto
        self.next_frame.pack_forget()
        self.welcome_frame.pack()


# Crea la finestra principale
app = MainWindow()
app.mainloop()