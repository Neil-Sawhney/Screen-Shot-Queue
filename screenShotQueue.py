import time
import tkinter
from tkinter import ttk
from pynput import mouse, keyboard
from PIL import Image
import mss
import glob
from io import BytesIO
import win32clipboard
import os
import sys

class stuff(object):
    presses = set()
    imgNum = 0
    leftArray = []
    sleepTime = 0.4

    def send_to_clipboard(self, imgNum):
        image = Image.open(".\\images\\im" + str(imgNum) + ".png")
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        while(True):
            try:
                win32clipboard.OpenClipboard()
                break;
            except:
                pass
        
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        while(True):
            try:
                if(win32clipboard.GetClipboardData(win32clipboard.CF_DIB)[:-1] == data):
                    break;
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            except Exception as e:
                pass

        win32clipboard.CloseClipboard()
 
    def grab(self, x1y1, x2y2):
        with mss.mss() as sct:
            monitor = {"top": x1y1[1], "left": x1y1[0], "width": x2y2[0] - x1y1[0], "height": x2y2[1] - x1y1[1]}
            output = ".\\images\\im" + str(self.imgNum) + ".png"

            sct_img = sct.grab(monitor)
            mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
        print("im" + str(self.imgNum) + " saved")
        self.imgNum += 1

    def on_move(self, x, y):
        self.x = x
        self.y = y
        
    def on_press(self, key):

        COMBINATION = { keyboard.Key.alt_l, keyboard.KeyCode.from_char('2') }

        # if we press alt + 2 then add a point
        if key in COMBINATION:
            self.presses.add(key)
            if all (k in self.presses for k in COMBINATION):
                self.leftArray.append([ self.x, self.y ])
                print(self.leftArray)
                if(len(self.leftArray) % 2 == 0):
                    self.grab(self.leftArray[2*self.imgNum], self.leftArray[2*self.imgNum + 1])

        # if we press alt + 1 then remove a point/img
        COMBINATION = { keyboard.Key.alt_l, keyboard.KeyCode.from_char('1') }
        if key in COMBINATION:
            self.presses.add(key)
            if all (k in self.presses for k in COMBINATION):
                #pop the last element in leftArray
                try:
                    if(len(self.leftArray) % 2 == 0):
                        print("im" + str(self.imgNum) + " removed")
                        self.imgNum -= 1

                except IndexError:
                    print("No more elements in array")
                self.leftArray.pop()
                print(self.leftArray)

        # if we press alt + 3 then print and exit
        COMBINATION = { keyboard.Key.alt_l, keyboard.KeyCode.from_char('3') }
        if key in COMBINATION:
            self.presses.add(key)
            if all (k in self.presses for k in COMBINATION):
                return False
        
    def on_release(self, key):
        try:
            self.presses.remove(key)
        except KeyError:
            pass


#remove all images
dir = ".\\images"
filelist = glob.glob(os.path.join(dir, "*.png"))
for f in filelist:
    os.remove(f)

s = stuff()

# Print instructions
print("Press alt + 1 to remove a point")
print("Press alt + 2 to add a point")
print("Press alt + 3 to paste all and exit")


with mouse.Listener(on_move=s.on_move) as listener:
    with keyboard.Listener(on_press=s.on_press, on_release=s.on_release) as listener:
        listener.join()

###################### DO STUFF ##################### 
def pasteAll():
    for i in range(s.imgNum):
        time.sleep(s.sleepTime)
        s.send_to_clipboard(i)
        time.sleep(s.sleepTime)

        #press ctrl v
        keyboard.Controller().press(keyboard.Key.ctrl)
        keyboard.Controller().press('v')
        keyboard.Controller().release(keyboard.Key.ctrl)
        keyboard.Controller().release('v')


# ask the user if they want to paste all images or just open the folder

popup = tkinter.Tk()

# make the popup appear on top of everything
popup.wm_attributes("-topmost", 1)

# resize 
popup.geometry("300x150")

# make the popup not resizable
popup.resizable(0,0)

# make the popup the active window
popup.focus_force()

Label = ttk.Label(popup, text="CHOOSE NOW")
Label.pack()
B1 = ttk.Button(popup, text="Dont Paste", command = lambda: [popup.destroy, os.startfile(".\\images"), sys.exit()])

# make B1 the default button
B1.pack()
popup.mainloop()


# Open the folder with the images
pasteAll()
os.startfile(".\\images")