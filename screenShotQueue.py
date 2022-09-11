import time
from pynput import mouse, keyboard
from PIL import Image
import mss
import glob
from io import BytesIO
import win32clipboard
import os

class stuff(object):
    presses = set()
    imgNum = 0
    leftArray = []
    sleepTime = 0.4


    def send_to_clipboard(self, imgNum):
        image = Image.open("C:\\Users\\neils\\OneDrive\\Documents\\Programming\\bots\\Screen Shot Queue\\images\\im" + str(imgNum) + ".png")
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()
 
    def grab(self, x1y1, x2y2):
        with mss.mss() as sct:
            monitor = {"top": x1y1[1], "left": x1y1[0], "width": x2y2[0] - x1y1[0], "height": x2y2[1] - x1y1[1]}
            output = "C:\\Users\\neils\\OneDrive\\Documents\\Programming\\bots\\Screen Shot Queue\\images\\im" + str(self.imgNum) + ".png"

            sct_img = sct.grab(monitor)
            mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
        print("im" + str(self.imgNum) + " saved")
        self.imgNum += 1

    def on_move(self, x, y):
        self.x = x
        self.y = y
        
    def on_press(self, key):

        self.presses.add(key)

        # if we press ctrl + right then add a point
        if self.presses ==  {keyboard.Key.ctrl_l, keyboard.Key.right}:
            self.leftArray.append([ self.x, self.y ])
            print(self.leftArray)
            if(len(self.leftArray) % 2 == 0):
                self.grab(self.leftArray[2*self.imgNum], self.leftArray[2*self.imgNum + 1])

        # if we press ctrl + left then remove a point/img
        if self.presses == {keyboard.Key.ctrl_r, keyboard.Key.left}:
            #pop the last element in leftArray
            try:

                if(len(self.leftArray) % 2 == 0):
                    self.leftArray.pop()
                    self.leftArray.pop()
                    print("im" + str(self.imgNum) + " removed")
                    self.imgNum -= 1
                else:
                    self.leftArray.pop()

            except IndexError:
                print("No more elements in array")
            print(self.leftArray)

        if key == keyboard.Key.esc:
            return False

    def on_release(self, key):
        try:
            self.presses.remove(key)
        except KeyError:
            pass

#remove all images
dir = "C:\\Users\\neils\\OneDrive\\Documents\\Programming\\bots\\Screen Shot Queue\\images"
filelist = glob.glob(os.path.join(dir, "*.png"))
for f in filelist:
    os.remove(f)

s = stuff()

with mouse.Listener(on_move=s.on_move) as listener:
    with keyboard.Listener(on_press=s.on_press, on_release=s.on_release) as listener:
        listener.join()

###################### DO STUFF ##################### 

for i in range(s.imgNum):
    s.send_to_clipboard(i)

    time.sleep(s.sleepTime)
    #press ctrl v
    keyboard.Controller().press(keyboard.Key.ctrl)
    keyboard.Controller().press('v')
    keyboard.Controller().release(keyboard.Key.ctrl)
    keyboard.Controller().release('v')
    time.sleep(s.sleepTime)
