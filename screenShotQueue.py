import time
import tkinter
from tkinter import ttk, filedialog
from pynput import mouse, keyboard
from PIL import Image
import mss
import glob
from io import BytesIO
import pywintypes
import win32clipboard
import os
import sys
import shutil
import nbformat as nbf

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
        print("im" + str(self.imgNum + 1) + " saved")
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
                # if this is the first point then delete all images
                if(len(self.leftArray) == 0):
                    self.removeImages()
                self.leftArray.append([ self.x, self.y ])
                print(self.leftArray)
                if(len(self.leftArray) % 2 == 0):
                    # if the second point is not down and to the right of the first point then flip them
                    if(self.leftArray[-2][0] > self.leftArray[-1][0] or self.leftArray[-2][1] > self.leftArray[-1][1]):
                        self.leftArray[-2], self.leftArray[-1] = self.leftArray[-1], self.leftArray[-2]
                    self.grab(self.leftArray[-2], self.leftArray[-1])

        # if we press alt + 1 then remove a point/img
        COMBINATION = { keyboard.Key.alt_l, keyboard.KeyCode.from_char('1') }
        if key in COMBINATION:
            self.presses.add(key)
            if all (k in self.presses for k in COMBINATION):
                if(len(self.leftArray) == 0):
                    print("no points to remove")
                    return
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
    def removeImages(self):
        dir = ".\\images"
        filelist = glob.glob(os.path.join(dir, "*.png"))
        for f in filelist:
            os.remove(f)


s = stuff()

# Print instructions
print("Press alt + 1 to remove a point")
print("Press alt + 2 to add a point")
print("Press alt + 3 when done")


with mouse.Listener(on_move=s.on_move) as listener:
    with keyboard.Listener(on_press=s.on_press, on_release=s.on_release) as listener:
        listener.join()

###################### DO STUFF ##################### 
def pasteAll(numOfEnters):
    for i in range(s.imgNum):
        time.sleep(s.sleepTime)
        s.send_to_clipboard(i)
        time.sleep(s.sleepTime)

        #press ctrl v
        keyboard.Controller().press(keyboard.Key.ctrl)
        keyboard.Controller().press('v')
        keyboard.Controller().release(keyboard.Key.ctrl)
        keyboard.Controller().release('v')

        for j in range(numOfEnters):
            time.sleep(s.sleepTime)
            keyboard.Controller().press(keyboard.Key.enter)
            keyboard.Controller().release(keyboard.Key.enter)

def jupyterNotebook():
    # ask for directory
    root = tkinter.Tk()
    root.withdraw()
    directory = filedialog.askdirectory()
    root.destroy()

    # create a new file
    nb = nbf.v4.new_notebook()

    # copy all images in .\images to the new directory
    files = os.listdir(".\\images")
    for f in files:
        shutil.copy(".\\images\\" + f, directory + "\\images\\" + f)
        #add json for this image to the file
        markdown = "![" + f + "](images\\" + f + ")"
        nb['cells'].append(nbf.v4.new_markdown_cell(markdown))

    nbf.write(nb, directory + "\\jupyterNotebook.ipynb")
    os.startfile(directory + "\\jupyterNotebook.ipynb")

def createPdf():
    images = [
        Image.open(".\\images\\im" + str(i) + ".png") for i in range(s.imgNum + 1)
    ]

    # ask for directory
    root = tkinter.Tk()
    root.withdraw()
    pdf_path = filedialog.askdirectory()
    root.destroy()

    # save the images to a pdf
    images[0].save(pdf_path + "\\NeilSawhney.pdf", save_all=True, append_images=images[1:])

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
B1 = ttk.Button(popup, text="Just Open Folder", command = lambda: [popup.destroy, os.startfile(".\\images"), sys.exit()])
B2 = ttk.Button(popup, text="Paste Normally", command = lambda: [popup.destroy, pasteAll(0), os.startfile(".\\images"), sys.exit()])
B3 = ttk.Button(popup, text="Paste with Enter", command = lambda: [popup.destroy, pasteAll(1), os.startfile(".\\images"), sys.exit()])
B4 = ttk.Button(popup, text="Paste with double Enter", command = lambda: [popup.destroy, pasteAll(2), os.startfile(".\\images"), sys.exit()])
B5 = ttk.Button(popup, text="Jupyter Notebook", command = lambda: [popup.destroy, jupyterNotebook(), sys.exit()])
B6 = ttk.Button(popup, text="Create PDF", command = lambda: [popup.destroy, createPdf(), sys.exit()])
B1.pack()
B2.pack()
B3.pack()
B4.pack()
B5.pack()
B6.pack()

popup.mainloop()

os.startfile(".\\images")