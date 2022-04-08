from modules.app import App
import threading

class MainApp:

    def __init__(self):
        self.__app = App()

    def run(self):
        self.__app.mainloop()



if __name__ == "__main__":
    app = MainApp()
    # app2 = MainApp()
    x = threading.Thread(target=app.run(), daemon=True)
    # y = threading.Thread(target=app2.run(), daemon=True)
    x.start()
