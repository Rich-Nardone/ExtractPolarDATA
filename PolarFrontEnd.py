import tkinter as tk
import PolarData as PD

def OnClick_PULL():
    files = PD.Start()
    btn.destroy()
    lbl.destroy()
    if(len(files.keys())>0):
        for key,val in files.items(): 
            newlbl = tk.Label(window, text="Excel File "+str(key)+" contains: "+str(val)+" Student Athletes").pack()
    else:
        newlbl = tk.Label(window, text="No Polar Data has been uploaded today")
window = tk.Tk()

window.title("Polar Data Extracter")

window.geometry('350x200')

lbl = tk.Label(window, text="Pull Today's Session Data:")
btn = tk.Button(window, text="Pull", command=lambda: OnClick_PULL())

lbl.pack()
btn.pack()
window.mainloop()
