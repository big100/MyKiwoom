import pyttsx3


class Sound:
    def __init__(self, qlist):
        self.soundQ = qlist[3]
        self.text2speak = pyttsx3.init()
        self.text2speak.setProperty('rate', 170)
        self.text2speak.setProperty('volume', 1.0)
        self.Start()

    def __del__(self):
        self.text2speak.stop()

    def Start(self):
        while True:
            text = self.soundQ.get()
            self.text2speak.say(text)
            self.text2speak.runAndWait()
