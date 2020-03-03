import xml.etree.ElementTree as ET

titles = ['Adobe InDesign CC 2018', 'Acrobat DC', 'Adobe Photoshop CC 2018', 'Adobe Illustrator CC 2018', 'Adobe Lightroom Classic CC', 'Adobe Lightroom CC']

def reset_trial(title):
    path = "C:/Program Files/Adobe/"
    end_path = "/AMT/application.xml"

    if title == 'Adobe Illustrator CC 2018':
        path = "C:/Program Files/Adobe/"
        end_path ="/Support Files/Contents/Windows/AMT/application.xml"

    if title == 'Acrobat DC':
        path = "C:/Program Files (x86)/Adobe/"
        end_path = "/Acrobat/AMT/application.xml"

    tree = ET.parse(path + title + end_path)
    root = tree.getroot()


    # Find TrialSerialNumber
    trial_key = root.find("./Other/Data[@key='TrialSerialNumber']")

    print(trial_key.text)

    # Add 10 to the TrialSerialNumber to reset
    trial_num = int(trial_key.text) + 10

    print(trial_num)

    # Convert TrialSerialNumber back to string
    trial_key.text = str(trial_num)

    print(trial_key.text)

    #tree.write("Data/application2.xml")
    tree.write(path + title + end_path)


for i in range(0,len(titles)):
    reset_trial(titles[i])
