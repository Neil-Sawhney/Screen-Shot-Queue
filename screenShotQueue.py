import base64
import glob
import os
import re
import shutil
import sys
import time
import tkinter
from io import BytesIO
from tkinter import filedialog, ttk

import mss
import nbformat as nbf
import pywintypes
import win32clipboard
from PIL import Image
from pynput import keyboard, mouse


class stuff(object):
    presses = set()
    imgNum = 0
    leftArray = []
    sleepTime = 0.4

    def send_to_clipboard(self, filePath):
        image = Image.open(filePath)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        while True:
            try:
                win32clipboard.OpenClipboard()
                break
            except:
                pass

        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        while True:
            try:
                if win32clipboard.GetClipboardData(win32clipboard.CF_DIB)[:-1] == data:
                    break
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            except Exception as e:
                pass

        win32clipboard.CloseClipboard()

    def grab(self, x1y1, x2y2):
        try:
            with mss.mss() as sct:
                monitor = {
                    "top": x1y1[1],
                    "left": x1y1[0],
                    "width": x2y2[0] - x1y1[0],
                    "height": x2y2[1] - x1y1[1],
                }

                # convert num to left padded string
                num = str(self.imgNum)
                num = num.zfill(4)

                output = ".\\images\\" + num + ".png"

                sct_img = sct.grab(monitor)
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
            print("im" + str(self.imgNum + 1) + " saved")
            self.imgNum += 1
        except Exception as e:
            return False

    def on_move(self, x, y):
        self.x = x
        self.y = y

    def on_press(self, key):
        COMBINATION = {keyboard.Key.alt_l, keyboard.KeyCode.from_char("2")}

        # if we press alt + 2 then add a point
        if key in COMBINATION and len(self.presses) < 2:
            self.presses.add(key)
            if all(k in self.presses for k in COMBINATION):
                # if this is the first point then delete all images
                if len(self.leftArray) == 0:
                    self.removeImages()

                self.leftArray.append([self.x, self.y])

                if len(self.leftArray) % 2 != 0:
                    print("\nstart point recorded")
                    self.toggleCrosshair()
                else:
                    print("end point recorded")
                    self.toggleCrosshair()

                    firstPointX = self.leftArray[-2][0]
                    secondPointX = self.leftArray[-1][0]
                    firstPointY = self.leftArray[-2][1]
                    secondPointY = self.leftArray[-1][1]

                    # if the second point is up and left, swap the points
                    if secondPointY < firstPointY and secondPointX < firstPointX:
                        self.leftArray[-2][0] = secondPointX
                        self.leftArray[-2][1] = secondPointY
                        self.leftArray[-1][0] = firstPointX
                        self.leftArray[-1][1] = firstPointY
                    # if the second point is up and right, swap the y values
                    elif secondPointY < firstPointY and secondPointX > firstPointX:
                        self.leftArray[-2][1] = secondPointY
                        self.leftArray[-1][1] = firstPointY
                    # if the second point is down and left, swap the x values
                    elif secondPointY > firstPointY and secondPointX < firstPointX:
                        self.leftArray[-2][0] = secondPointX
                        self.leftArray[-1][0] = firstPointX

                    # if the grab fails, give up
                    if (self.grab(self.leftArray[-2], self.leftArray[-1])) == False:
                        print("An unknown error occured, whoops")
                        return False

        # if we press alt + 1 then remove a point/img
        COMBINATION = {keyboard.Key.alt_l, keyboard.KeyCode.from_char("1")}
        if key in COMBINATION and len(self.presses) < 2:
            self.presses.add(key)
            if all(k in self.presses for k in COMBINATION):
                if len(self.leftArray) == 0:
                    print("\nno points to remove")
                    return

                # pop the last element in leftArray
                self.leftArray.pop()
                if len(self.leftArray) % 2 != 0:
                    self.removeLastImage()

                    print("\nend point removed")
                    print("im" + str(self.imgNum + 1) + " removed")
                else:
                    print("start point removed")

        # if we press alt + 3 then print and exit
        COMBINATION = {keyboard.Key.alt_l, keyboard.KeyCode.from_char("3")}
        if key in COMBINATION and len(self.presses) < 2:
            self.presses.add(key)
            if all(k in self.presses for k in COMBINATION):
                return False

    def on_release(self, key):
        try:
            self.presses.remove(key)
        except KeyError:
            pass

    # remove all images
    def removeImages(self):
        dir = ".\\images"
        filelist = glob.glob(os.path.join(dir, "*.png"))
        for f in filelist:
            os.remove(f)

    def removeLastImage(self):
        self.imgNum -= 1
        dir = ".\\images"
        filelist = glob.glob(os.path.join(dir, "*.png"))
        os.remove(filelist[-1])

    def toggleCrosshair(self):
        # send win + alt + p
        keyboard.Controller().press(keyboard.Key.cmd)
        keyboard.Controller().press(keyboard.Key.alt_l)
        keyboard.Controller().press("p")

        keyboard.Controller().release(keyboard.Key.cmd)
        keyboard.Controller().release(keyboard.Key.alt_l)
        keyboard.Controller().release("p")


s = stuff()

try:
    newDir = ".\\images"
    os.makedirs(newDir)
except OSError:
    pass

# Print instructions
print("Press alt + 1 to remove a point")
print("Press alt + 2 to add a point")
print("Press alt + 3 when done")


with mouse.Listener(on_move=s.on_move) as listener:
    with keyboard.Listener(on_press=s.on_press, on_release=s.on_release) as listener:
        listener.join()


###################### DO STUFF #####################
def pasteAll(numOfEnters, popup):
    popup.destroy()

    # find all images in .\images
    files = os.listdir(".\\images")

    # sort the images by name
    files.sort(key=lambda f: int(re.sub("\D", "", f)))

    for f in files:
        time.sleep(s.sleepTime)
        s.send_to_clipboard(".\\images\\" + f)
        time.sleep(s.sleepTime)

        # press ctrl v
        keyboard.Controller().press(keyboard.Key.ctrl)
        keyboard.Controller().press("v")
        keyboard.Controller().release(keyboard.Key.ctrl)
        keyboard.Controller().release("v")

        for j in range(numOfEnters):
            time.sleep(s.sleepTime)
            keyboard.Controller().press(keyboard.Key.enter)
            keyboard.Controller().release(keyboard.Key.enter)


def jupyterNotebook():
    root = tkinter.Tk()
    root.withdraw()
    # save as file name with default name output.pdf
    pdf_path = filedialog.asksaveasfilename(
        defaultextension=".ipynb", initialfile="jupyterNotebook.ipynb"
    )
    root.destroy()

    pdf_directory = os.path.dirname(pdf_path)

    # create a new file
    nb = nbf.v4.new_notebook()

    # copy all images in .\images to the new directory
    files = os.listdir(".\\images")

    # sort the images by name
    files.sort(key=lambda f: int(re.sub("\D", "", f)))

    for f in files:
        with open(".\\images\\" + f, "rb") as f:
            encoded_image = base64.b64encode(f.read()).decode("utf-8")

        markdown = "![image.png](attachment:image.png)"

        cell = nbf.v4.new_markdown_cell(markdown)
        cell["attachments"] = {"image.png": {"image/png": encoded_image}}

        nb["cells"].append(cell)

    nbf.write(nb, pdf_path)
    os.startfile(pdf_directory)


def createPdf():
    # ask for directory
    root = tkinter.Tk()
    root.withdraw()
    # save as file name with default name output.pdf
    pdf_path = filedialog.asksaveasfilename(
        defaultextension=".pdf", initialfile="output.pdf"
    )
    root.destroy()

    pdf_directory = os.path.dirname(pdf_path)

    # get the images from the images folder
    files = os.listdir(".\\images")
    # sort the images by name
    files.sort(key=lambda f: int(re.sub("\D", "", f)))

    # create a new pdf using PIL
    im1 = Image.open(".\\images\\" + files[0])
    im1.save(
        pdf_path,
        "PDF",
        resolution=100.0,
        save_all=True,
        append_images=[Image.open(".\\images\\" + f) for f in files[1:]],
    )

    os.startfile(pdf_directory)


# ask the user if they want to paste all images or just open the folder
popup = tkinter.Tk()
# make the popup appear on top of everything
popup.wm_attributes("-topmost", 1)
# make the popup the active window
popup.focus_force()
# make the popup appear in the top right corner
popup.geometry("+%d+%d" % (popup.winfo_screenwidth() - 350, 0))

Label = ttk.Label(popup, text="CHOOSE NOW")
Label.pack()
B1 = ttk.Button(
    popup,
    text="Just Open Folder",
    command=lambda: [popup.destroy, os.startfile(".\\images"), sys.exit()],
)
B2 = ttk.Button(
    popup,
    text="Paste Normally",
    command=lambda: [
        popup.destroy,
        pasteAll(0, popup),
        os.startfile(".\\images"),
        sys.exit(),
    ],
)
B3 = ttk.Button(
    popup,
    text="Paste with Enter",
    command=lambda: [
        popup.destroy,
        pasteAll(1, popup),
        os.startfile(".\\images"),
        sys.exit(),
    ],
)
B4 = ttk.Button(
    popup,
    text="Paste with double Enter",
    command=lambda: [
        popup.destroy,
        pasteAll(2, popup),
        os.startfile(".\\images"),
        sys.exit(),
    ],
)
B5 = ttk.Button(
    popup,
    text="Jupyter Notebook",
    command=lambda: [popup.destroy, jupyterNotebook(), sys.exit()],
)
B6 = ttk.Button(
    popup, text="Create PDF", command=lambda: [popup.destroy, createPdf(), sys.exit()]
)
B1.pack()
B2.pack()
B3.pack()
B4.pack()
B5.pack()
B6.pack()

popup.mainloop()
