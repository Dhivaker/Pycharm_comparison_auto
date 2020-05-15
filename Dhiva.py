import os
import shutil
import numpy as np
import matplotlib.pyplot as plt
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
    for filename in os.listdir(New_results_dir):
        if filename.contains('coss'):
            return{"x_label":"Time/uSecs","y_label":"Coss/pF","lengend":"Coss","title":"Output capacitance","xmin":"0","ymin":"0","xmax":"2","ymax":"10","xunit":"0.2","yunit":"2"}
        if filename.contains("output_25C"):
            return{"x_label":"VCE/V","y_label":"IC/A","lengend":"Output TJ=25C","title":"Output Characteristics","xmin":"0","ymin":"0","xmax":"10","ymax":"20","xunit":"0.5","yunit":"2.5"}

def generate_plot():
    for files in os.listdir(New_results_dir):

        if files.endswith('.txt'):
            print(files)
            x, y = [], []
            with open(os.path.join(New_results_dir, files)) as txt_file:
                     raw_data = txt_file.readlines()
                     if (len(raw_data[0].split())> 2):
                         index_x = 0
                         index_y = 2
                     else:
                         index_x = 0
                         index_y = 1
                     labels = get_labels(files)


                     for itr in range(1, len(raw_data)):
                         x.append(float(raw_data[itr].split()[index_x].strip()))
                         y.append(float(raw_data[itr].split()[index_y].strip()))
                     plt.plot(x, y, label=labels["lengend"])
                     plt.xlabel(labels["x_label"])
                     plt.ylabel(labels["y_label"])
                     plt.title(labels["title"])
                     plt.legend()


                     #plt.show()
                     image_file = files.replace('data','image').replace('txt','png')
                     #plt.savefig("Dhiva.png")
                     plt.savefig(os.path.join(New_results_dir, image_file))
                     plt.clf()
#create_folder()
#generate_netfiles()

generate_plot()
get_labels(filename)

