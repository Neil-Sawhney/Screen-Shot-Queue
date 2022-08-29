import time
from pynput import mouse, keyboard
from PIL import Image
import mss
import glob
from io import BytesIO
import win32clipboard
import os

class stuff(object):
    imgNum = 0

    leftArray = []
    sleepTime = 0.1


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

    def on_click(self, x, y, button, pressed):
        # if left click, add coordinates to array
        if button == mouse.Button.middle:
            if not(pressed):
                self.leftArray.append([ x, y ])
                print(self.leftArray)
                if(len(self.leftArray) % 2 == 0):
                    self.grab(self.leftArray[2*self.imgNum], self.leftArray[2*self.imgNum + 1])
        if button == mouse.Button.right:
            if not(pressed):
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

        if button == mouse.Button.left:
            if not(pressed) and self.leftArray != []:
                return False

    def on_press(self, key):
        if key == keyboard.Key.esc:
            return False

#remove all images
dir = "C:\\Users\\neils\\OneDrive\\Documents\\Programming\\bots\\Screen Shot Queue\\images"
filelist = glob.glob(os.path.join(dir, "*.png"))
for f in filelist:
    os.remove(f)

s = stuff()
with mouse.Listener(
        on_click=s.on_click,
        ) as listener:
    listener.join()

with keyboard.Listener(
        on_press=s.on_press,
) as listener:
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
