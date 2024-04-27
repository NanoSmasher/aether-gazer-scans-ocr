import win32gui, win32com.client
import pandas as pd
from PIL import Image
import PIL.ImageGrab
import pytesseract
from pathlib import Path
from openpyxl import load_workbook
### for colour recognition
import cv2 as cv
import numpy as np
import math

### Variables
# Path to tesseract binary. Instructions: https://github.com/tesseract-ocr/tesseract#installing-tesseract
tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = tesseract_cmd # todo: error check
# pytesseract.get_tesseract_version()
# Version('5.3.1.20230401')
game_title = 'aether gazer'
filename = 'Aether Gazer Pulls.xlsx'

def create_file():
    # create pulls file if it doesn't exist
    if not Path(filename).is_file():
        print(f"{filename} not found, creating...")
        with pd.ExcelWriter(filename) as writer:
            empty_pd = pd.DataFrame({}, columns=['Scan Time','Type','Name','Rarity','A Counter','S Counter'])
            empty_pd.to_excel(writer, sheet_name='Limited', header=True, index=False)  
            empty_pd.to_excel(writer, sheet_name='Standard', header=True, index=False)  
            empty_pd.to_excel(writer, sheet_name='Functor', header=True, index=False) 

#Screenshot Method
def image_from_screenshot(): #game_title = 'aether gazer'
    toplist, winlist = [], [] 
    # create list of all windows (Windows only)
    def enum_cb(hwnd, results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
    win32gui.EnumWindows(enum_cb, toplist)
    # grab the hwnd for active window and first window matching Aether Gazer (Bluestacks only?)
    aw = win32gui.GetForegroundWindow() # this should be the active terminal
    ag = [(hwnd, title) for hwnd, title in winlist if game_title in title.lower()]
    # get image from that first window result
    # todo: error check
    hwnd = ag[0][0]
    # bring window to foreground and snap
    win32com.client.Dispatch("WScript.Shell").SendKeys('%') #prevents error: (0, 'SetForegroundWindow', 'No error message is available')
    win32gui.SetForegroundWindow(hwnd)
    bbox = win32gui.GetWindowRect(hwnd)
    screenshot = PIL.ImageGrab.grab(bbox)
    imgnp = np.array(screenshot) 
    imgcvt = cv.cvtColor(imgnp, cv.COLOR_BGR2RGB) # convert colour is necessary as PIL and numpy use different format
    h, w, _ = imgcvt.shape
    cropped = imgcvt[int(0.245*h):int(0.8*h),int(0.1*w):int(0.85*w)]
    win32com.client.Dispatch("WScript.Shell").SendKeys('%') #prevents error: (0, 'SetForegroundWindow', 'No error message is available')
    win32gui.SetForegroundWindow(aw) # return to window
    return cropped

#File Method
def image_from_file():
    f='test/4jpeg.png'
    pImage = Image.open(f)
    w, h = pImage.size 
    cropped = pImage.crop((0.1*w, 0.175*h, 0.9*w, 0.8*h)) # left,top,right,bottom
    return cropped         

def data_text_box(d):
    text = [''] * (max(d['block_num']) + 1)
    box = [None] * (max(d['block_num']) + 1)
    for i in range(len(d['level'])):
        if d['par_num'][i] == 0:
            (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
            box[d['block_num'][i]] = (y, y + h, x, x + w)
            #print(f"block #{d['block_num'][i]} found at {x}:{x+w}, {y}:{y+h}")
            #_ = cv.rectangle(imgc, (x, y), (x + w, y + h), (0, 255, 0), 2) #draw box on original image
        elif d['conf'][i] != -1:
            text[d['block_num'][i]] += d['text'][i] + " "
            #print(f"{d['text'][i]}")
    trimmed = text
    for e,v in enumerate(text):
        if len(v) > 0:
            trimmed[e] = v[:-1]
    return trimmed, box

#grab the most dominant colour
def dominant_colour(img,n):
    data = np.reshape(img, (-1,3))
    data = np.float32(data)
    criteria = (cv.TERM_CRITERIA_EPS + cv.TERM_CRITERIA_MAX_ITER, 10, 1.0)
    flags = cv.KMEANS_RANDOM_CENTERS
    #compactness,labels,centers cv.kmeans(NumpyArray,K clusters,labels,attempts before giving up,flags)
    _,_,centers = cv.kmeans(data,n,None,criteria,10,flags)
    return centers #most to least dominant

#figure out text colour of tuple as a rarity number
def get_colour(col):
    #(b,g,r) tuples of white, black, blue, purple, gold
    colour_list = [(220, 220, 220),(78, 78, 78),(134, 88, 42),(137, 92, 161),(140, 170, 203)] 
    nearest = min(colour_list, key=lambda c: math.hypot(c[0] - col[0], c[1] - col[1], c[2] - col[2])) # magic
    if nearest == (220, 220, 220): return 0 #'white'
    if nearest == (78, 78, 78): return 3 #'black'
    if nearest == (134, 88, 42): return 3 #'blue' #black and blue are the same rarity
    if nearest == (137, 92, 161): return 4 #'purple'
    if nearest == (140, 170, 203): return 5 #'gold'
    return -1

def extract_screenshot():
    im = image_from_screenshot()
    h, w, _ = im.shape # height, width, ??? of image
    imType = im[:,0:int(0.15*w)]
    imName = im[:,int(0.15*w):int(0.75*w)]
    imTime = im[:,int(0.75*w):]

    #get colour (aka. rarity) data from screenshot
    colourData = pytesseract.image_to_data(imType, config='--psm 11', output_type=pytesseract.Output.DICT) # Page Segmentation Modes, default is 3 which is fully automatic, 11 is sparse text
    t,b = data_text_box(colourData)
    datRare = []
    for i in range(1,len(t)): # index 0 is the entire image so we skip
        roi = im[b[i][0]:b[i][1],b[i][2]:b[i][3]]
        centers = dominant_colour(roi,2)
        textColour = max(get_colour(centers[0]),get_colour(centers[1]))
        datRare.append(textColour)

    # Get the normal data and end up with table
    datType = list(filter(None,pytesseract.image_to_string(imType, config='--psm 11').split('\n')))
    datName = list(filter(None,pytesseract.image_to_string(imName, config='--psm 11').split('\n')))
    datTime = list(filter(None,pytesseract.image_to_string(imTime, config='--psm 11').split('\n')))

    df = pd.DataFrame(list(zip(datTime, datType, datName, datRare)),columns =['Scan Time','Type','Name','Rarity'])
    df = df.reindex(index=df.index[::-1]) #flip
    df['A Counter'] = '' # create new empty columns (otherwise it will fill up with NaNs and change counter to a float)
    df['S Counter'] = ''

    return df

def main():
    create_file() # create new excel file if necessary
    df = pd.read_excel(filename,sheet_name=['Standard','Limited','Functor']) #dtype={'A Counter': int,'S Counter': int} messing up...
    complete_df = df
    print("File loaded")

    close = False
    while not close:
        print("Choose which sheet to save entries into.")
        i = input("S for Standard, L for Limited, F for Functor, (X) to close: ")
        i = i.lower()
        if (i == 's') or (i == 'l') or (i == 'f'):
            if i == 's': sname = 'Standard'
            elif i == 'l': sname = 'Limited'
            elif i == 'f': sname = 'Functor'
            
            close2 = False            
            while not close2:
                entries = len(df[sname].index)
                if entries > 0:
                    print("=====Last 10 Entries=====")
                    print(df[sname][-10:])
                else:
                    print(f"{sname} Sheet Currently Emtpy")

                input("Ready. Press Enter to extract screenshot.")
                
                edf = extract_screenshot()
                print(f"\n=====Extracted Entries=====\n")
                print(edf[-10:])

                i2 = int(input(f"\nHow many entries to add from extracted screenshot? Enter number from 1 to {len(edf.index)}, (0) to quit: ") or "0")
                if (i2 > 0) and (i2 <= len(edf.index)):
                    i2 = -i2
                    df[sname] = pd.concat([df[sname],edf[i2:]], ignore_index=True)

                    #update the A and S counters
                    for entry in range(entries,len(df[sname].index)):
                        if entry == 0: #blank sheet
                            df[sname].loc[entry,'A Counter'] = 1
                            df[sname].loc[entry,'S Counter'] = 1
                            continue
                        rarity = int(df[sname].loc[entry - 1,'Rarity'])
                        if rarity == 5:
                            df[sname].loc[entry,'S Counter'] = 1
                        else:
                            df[sname].loc[entry,'S Counter'] = df[sname].loc[entry - 1,'S Counter'] + 1
                        if rarity == 4:
                            df[sname].loc[entry,'A Counter'] = 1
                        else:
                           df[sname].loc[entry,'A Counter'] = df[sname].loc[entry - 1,'A Counter'] + 1

                    print("Saving entries...\n\n\n")
                    with pd.ExcelWriter(filename,mode="a",if_sheet_exists="replace") as writer:
                        df['Limited'].to_excel(writer, sheet_name="Limited", header=True, index=False)
                        df['Standard'].to_excel(writer, sheet_name="Standard", header=True, index=False)
                        df['Functor'].to_excel(writer, sheet_name="Functor", header=True, index=False)
                    
                else:
                    close2 = True


        else:
            print("Closing program...")
            close = True


            
if __name__=="__main__":
    main()