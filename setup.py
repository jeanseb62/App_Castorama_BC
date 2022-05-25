print("Chargement en cours...")

import subprocess

subprocess.call(['pip', 'install', 'numpy'])
subprocess.call(['pip', 'install', 'pandas'])
subprocess.call(['pip', 'install', 'openpyxl'])
subprocess.call(['pip', 'install', 'PyPDF2'])
subprocess.call(['pip', 'install', 'tabula-py'])
subprocess.call(['pip', 'install', 'glob2'])

print("Les modules ont bien été importés !")