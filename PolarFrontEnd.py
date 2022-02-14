import tkinter as tk
import PolarData as PD

def OnClick_PULL():
    files = PD.Start()
    if(len(files.keys())>0):
        for key,val in files.items(): 
            newlbl = tk.Label(window, text="Excel File "+str(key)+" contains: "+str(val)+" Student Athletes").pack()
    else:
        newlbl = tk.Label(window, text="No Polar Data has been uploaded today").pack()

def OnClick_Create():
    PD.CreateData(e1.get(),e2.get(),e3.get(),e4.get())
    lbl1.destroy()
    e1.destroy()
    lbl2.destroy()
    e2.destroy()
    lbl3.destroy()
    e3.destroy()
    lbl4.destroy()
    e4.destroy()
    Submitbtn.destroy()
    lbl = tk.Label(window, text="Pull Today's Session Data:").pack()
    btn = tk.Button(window, text="Pull", command=lambda: OnClick_PULL()).pack()


window = tk.Tk()

window.title("Polar Data Extracter")

window.geometry('350x200')

if(PD.CheckForSetup()):
    lbl = tk.Label(window, text="Pull Today's Session Data:").pack()
    btn = tk.Button(window, text="Pull", command=lambda: OnClick_PULL()).pack()
else:
    lbl1 = tk.Label(window, text="Polar Email:")  
    e1 = tk.Entry(window,textvariable=tk.StringVar())
    lbl2 = tk.Label(window, text="Polar Password:")
    e2 = tk.Entry(window,textvariable=tk.StringVar())
    lbl3 = tk.Label(window, text="Download Path:")
    e3 = tk.Entry(window,textvariable=tk.StringVar())
    lbl4 = tk.Label(window, text="Polar Path:")
    e4 = tk.Entry(window,textvariable=tk.StringVar())
    lbl1.pack()
    e1.pack()
    lbl2.pack()
    e2.pack()
    lbl3.pack()
    e3.pack()
    lbl4.pack()
    e4.pack()
    Submitbtn = tk.Button(window, text="Submit", command=lambda: OnClick_Create()).pack()
    

window.mainloop()
