import pyttsx3


class Sound:
    def __init__(self, qlist):
        """
           0        1        2      3      4       5       6      7       8        9       10       11
        windowQ, traderQ, receivQ, stgQ, soundQ, queryQ, teleQ, hoga1Q, hoga2Q, chart1Q, chart2Q, chart3Q,
        chart4Q, chart5Q, chart6Q, chart7Q, chart8Q, chart9Q, chart10Q, tick1Q, tick2Q, tick3Q, tick4Q
          12       13       14       15       16       17       18        19      20      21      22
        """
        self.soundQ = qlist[4]
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
