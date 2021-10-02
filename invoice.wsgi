import sys
sys.path.insert(0,'/home/attu/Desktop/invoiceGenerator')
activate_this = '/home/attu/Desktop/ScratchNest/env/bin/activate_this.py'
with open(activate_this) as file_:
    exec(file_.read(), dict(__file__=activate_this))

from invoiceApp.invoice import app as application