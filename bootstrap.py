import os, sys 
if getattr(sys, 'frozen', False): 
    os.chdir(os.path.dirname(sys.executable)) 
else: 
    os.chdir(os.path.dirname(os.path.abspath(__file__))) 
from main import PDFExtractorGUI 
if __name__ == "__main__": 
    app = PDFExtractorGUI() 
    app.root.mainloop() 
