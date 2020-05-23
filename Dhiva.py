import os
import shutil
import numpy as np
import matplotlib.pyplot as plt
import PyPDF2

from shutil import copyfile

dir = os.getcwd()
netlist_dir =os.path.join(dir,"Netlist")

def create_folder():
    if os.path.exists(os.path.join(dir,"Results")):
         print('Results folder already exists and overwriting')
         shutil.rmtree(os.path.join(dir,"Results"))
    os.makedirs('Results',mode=0o777)
New_results_dir =os.path.join(dir,"Results")
input_file1 = os.path.join(dir,"User_input.txt")
# #


def generate_netfiles():
    with open(input_file1) as input:
        input_file = input.readlines()
    Location1 = input_file[4].split('=')[1].strip()
    Location2 = input_file[5].split('=')[1].strip()
    IGBT_Name = input_file[11].split('=')[1].strip()
    Diode_Name = input_file[12].split('=')[1].strip()

    for files in os.listdir(netlist_dir):
          if files.endswith('.net'):
              with open(os.path.join(netlist_dir, files)) as net_file:
                  net_file_lines = net_file.readlines()
    # #
                  for itr in range(0,len(net_file_lines)):
                      if "<<LocationIGBT>>" in net_file_lines[itr]:
                          net_file_lines[itr]=net_file_lines[itr].replace("<<LocationIGBT>>", Location1)
                      if "<<LocationDIODE>>" in net_file_lines[itr]:
                          net_file_lines[itr]=net_file_lines[itr].replace("<<LocationDIODE>>", Location2)
                      if "<<IGBT_Name>>" in net_file_lines[itr]:
                          net_file_lines[itr]=net_file_lines[itr].replace('<<IGBT_Name>>', IGBT_Name)
                      if "<<Diode_Name>>" in net_file_lines[itr]:
                          net_file_lines[itr]=net_file_lines[itr].replace("<<Diode_Name>>", Diode_Name)
              with open(os.path.join(New_results_dir, files), 'w')as output_file:
                   output_file.writelines(net_file_lines)
    # #
    # #     #print(net_file_lines)
    print('Started running simulations......Do some meditation')
    os.system("sim2 "+os.path.join(netlist_dir, "Script_all_simulations.sxscr"))
    print('Fininshed running simulations')
     # print(os.path.join(netlist_dir,"Script_all_simulations.sxscr"))
def get_labels(filename):

     if "coss" in filename:
         return{"x_label":"Time/Secs","y_label":"Cies, Crss, Coss/F","lengend":"Coss","title":"Capacitances","xmin":"0","ymin":10E-12,"xmax":"2","ymax":10E-8,"color":"b","width":2,"scale": "log"}
     if "output_25C" in filename:
         return{"x_label":"VCE/V","y_label":"IC/A","lengend":"Output TJ=25C","title":"Output Characteristics","color":"b","width":2,"scale": "linear", "xvalue": 1.6, "yvalue": 10.14, "lcolor": "k", "lstyle": "--"}
     if "output_175C" in filename:
         return{"x_label":"VCE/V","y_label":"IC/A","lengend":"Output TJ=175C","title":"Output Characteristics","color":"r","width":2,"scale": "linear", "xvalue": 1.85, "yvalue": 10.10, "lcolor": "k", "lstyle": "--"}
     if "transfer_25C" in filename:
         return{"x_label":"VGE/V","y_label":"IC/A","lengend":"Transfer TJ=25C","title":"Transfer Characteristics","xmin":"0","ymin":"0","xmax":"10","ymax":"30","xunit":"0.5","yunit":"2.5","color":"b","width":"2","scale": "linear","xvalue": 7.55, "yvalue": 0.01, "tx": 2, "ty": 4,"text":"VGE(th)@25°C = 7.55V"}
     if "transfer_175C" in filename:
         return{"x_label":"VGE/V","y_label":"IC/A","lengend":"Transfer TJ=175C","title":"Transfer Characteristics","xmin":"0","ymin":"0","xmax":"10","ymax":"30","xunit":"0.5","yunit":"2.5","color":"r","width":"2","scale": "linear"}
     if "vf_25C" in filename:
         return{"x_label":"VF/V","y_label":"IF/A","lengend":"Diode_Vf TJ=25C","title":"Diode forward Characteristics","xmin":"0","ymin":"0","xmax":"10","ymax":"35","xunit":"0.5","yunit":"2.5","color":"b","width":"2","scale": "linear","xvalue": 1.75, "yvalue": 11.59, "lcolor": "k", "lstyle": "--"}
     if "vf_175C" in filename:
         return{"x_label":"VF/V","y_label":"IF/A","lengend":"Diode_Vf TJ=175C","title":"Diode forward Characteristics","xmin":"0","ymin":"0","xmax":"10","ymax":"35","xunit":"0.5","yunit":"2.5","color":"r","width":"2","scale": "linear","xvalue": 1.5, "yvalue": 11.32, "lcolor": "k", "lstyle": "--"}
     if "crss" in filename:
         return{"x_label":"Time/Secs","y_label":"Cies, Crss, Coss/F","lengend":"Crss","title":"Capacitances","xmin":"0","ymin":10E-12,"xmax":"2","ymax":10E-8,"xunit":"0.2","yunit":"2","color":"g","width":"2","scale": "log","xvalue": 1.75, "yvalue": 11.59, "tx": "k", "ty": "--","text":"Crss = "}
     if "cies" in filename:
         return{"x_label":"Time/Secs","y_label":"Cies, Crss, Coss//pF","lengend":"Cies","title":"Capacitances","xmin":"0","ymin":10E-12,"xmax":"2","ymax":10E-8,"xunit":"0.2","yunit":"2","color":"r","width":"2","scale": "log","xvalue": 1.75, "yvalue": 11.59, "tx": "k", "ty": "--","text":"Cies = "}

def generate_plot():
    previous_file = ""
    for files in sorted(os.listdir(New_results_dir)):

        if files.endswith('.txt'):
            print(files)
            x, y = [], []

            with open(os.path.join(New_results_dir, files)) as txt_file:
                     raw_data = txt_file.readlines()
                     if (len(raw_data[0].split())> 2):
                         index_x = 1
                         index_y = 2
                     else:
                         index_x = 0
                         index_y = 1
                     labels = get_labels(files)
                     if previous_file != "" and previous_file.split('_')[0] != files.split('_')[0]:
                         image_file = previous_file.split('_')[0]+".png"
                         plt.savefig(os.path.join(New_results_dir, image_file), dpi= 300)
                         plt.clf()
                     previous_file = files
                     for itr in range(1, len(raw_data)):
                         x.append(float(raw_data[itr].split()[index_x].strip()))
                         y.append(float(raw_data[itr].split()[index_y].strip()))
                     plt.plot(x, y, label=labels["lengend"], color=labels["color"],linewidth=labels["width"], )
                     plt.grid(True)
                     plt.xlabel(labels["x_label"])
                     plt.ylabel(labels["y_label"])
                     plt.yscale(labels["scale"])
                     plt.title(labels["title"])
                     plt.legend()
                     if "output" in files or "diode" in files:
                         plt.axhline(labels["yvalue"],color=labels["lcolor"],linestyle=labels["lstyle"])
                         plt.axvline(labels["xvalue"],color=labels["lcolor"],linestyle=labels["lstyle"])
                     if "transfer_25C" in files:
                         plt.annotate(labels["text"], xytext=(labels["tx"], labels["ty"]), xy=(labels["xvalue"], labels["yvalue"]),arrowprops=dict(facecolor='black', shrink=0.05))
                     if "data" in files:
                         #plt.yticks(np.arange(10E-14, 10E-9, 10E-1/2))
                         plt.ylim(10E-14, 10E-9)
                         #plt.ylim(labels["ymin"], labels["ymax"])

    image_file = previous_file.split('_')[0] + ".png"
    plt.savefig(os.path.join(New_results_dir, image_file))
    plt.clf()

def pdfread():

    with open(os.path.join(dir,"IGBT7_Datasheet.pdf"),'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        # for itr in range(0, pdf_reader.numPages):
        #     print(pdf_reader.getPage(itr).extractText())
        print(pdf_reader.getPage(3).extractText())
    # igbt_datasheet = camelot.read_pdf(os.path.join(dir,"IGBT7_Datasheet.pdf"))
    # type(igbt_datasheet)






    # datasheet_read = PyPDF2.PdfFileReader(igbt_datasheet)
    # page = datasheet_read.getPage(0)
    # pagecontent = page.extractText()


#create_folder()
#generate_netfiles()

#pdfread()
generate_plot()



