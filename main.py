import tkinter as tk
import ui.ui as ui


def main():
    root = tk.Tk()
    root.title("Demo Processor")
    root.geometry("410x345")
    root.config(pady=10, padx=20)

    menubar = ui.DemoMenu(root)
    root.config(menu=menubar)

    frame = ui.DemoFrame(root)
    frame.pack()
    root.mainloop()


if __name__ == "__main__":
    main()
