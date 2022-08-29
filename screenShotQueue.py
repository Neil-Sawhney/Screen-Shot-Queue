import time
from pynput import mouse
from pynput.keyboard import Key, Controller
import pyscreenshot as ImageGrab
from PIL import Image
 
from io import BytesIO
import win32clipboard

class stuff(object):
    startPoint = []
    leftArray = []
    scrollArray = []
    sleepTime = 0.1

    def on_move(self, x, y):
        print('Pointer moved to {0}'.format(
            (x, y)))

    def on_click(self, x, y, button, pressed):
        # if left click, add coordinates to array
        if not(pressed):
            if self.startPoint == []:
                self.startPoint = [x,y]
                print("start point set to {0}".format(self.startPoint))
                return True
        if button == mouse.Button.left:
            if not(pressed):
                self.leftArray.append([ x, y ])
                print(self.leftArray)
                if(len(self.leftArray) % 2 == 0):
                    self.scrollArray.append(0)
                    print(self.scrollArray)
        if button == mouse.Button.right:
            if not(pressed):
                if self.leftArray == []:
                    self.startPoint = []
                    print("start point reset")
                    return True
                #pop the last element in leftArray
                try:
                    self.leftArray.pop()

                    if(len(self.leftArray) % 2 == 0):
                        self.scrollArray[-2] += self.scrollArray[-1]
                        self.scrollArray.pop()
                except IndexError:
                    print("No more elements in array")
                print(self.leftArray)

                if(len(self.leftArray) % 2 == 0):
                    print(self.scrollArray)
        # if middle mouse click
        if button == mouse.Button.middle:
            if not(pressed):
                return False

    def on_scroll(self, x, y, dx, dy):
        if(len(self.leftArray) % 2 == 0):
            self.scrollArray[-1] += dy
            print(self.scrollArray)

    def send_to_clipboard(self, clip_type, data):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()
 
    def grab(self, x1y1, x2y2):
        im = ImageGrab.grab(bbox=(x1y1[0], x1y1[1], x2y2[0], x2y2[1]))
        im.save("C:\\Users\\neils\\OneDrive\\Documents\\Programming\\bots\\Screen Shot Queue\\images\\im.png")
        image = Image.open("C:\\Users\\neils\\OneDrive\\Documents\\Programming\\bots\\Screen Shot Queue\\images\\im.png")
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        self.send_to_clipboard(win32clipboard.CF_DIB, data)

    #move mouse to first point
    def clickOrigin(self, currentIndex):
        for i in range(currentIndex):
            mouse.Controller().scroll(dy = -s.scrollArray[i], dx =0)
            time.sleep(s.sleepTime)

        #left click
        mouse.Controller().position = self.startPoint
        mouse.Controller().click(mouse.Button.left)

    def scrollToLocation(self, currentIndex):
        for i in range(currentIndex):
            mouse.Controller().scroll(dy = s.scrollArray[i], dx =0)
            time.sleep(s.sleepTime)


s = stuff()
with mouse.Listener(
        on_click=s.on_click,
        on_scroll=s.on_scroll) as listener:
    listener.join()

###################### DO STUFF ##################### 

keyboard= Controller()

#scroll back to the top
for i in range(len(s.scrollArray)):
    mouse.Controller().scroll(dy = -s.scrollArray[i], dx =0)
    time.sleep(s.sleepTime)


for i in range(0, len(s.leftArray), 2):

    s.grab(s.leftArray[i], s.leftArray[i+1])
    s.clickOrigin(i)
    time.sleep(s.sleepTime)

    #press ctrl v
    keyboard.press(Key.ctrl)
    keyboard.press('v')
    keyboard.release(Key.ctrl)
    keyboard.release('v')
    time.sleep(s.sleepTime)
