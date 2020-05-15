import os
import shutil
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from shutil import copyfile

dir = os.getcwd()
netlist_dir =os.path.join(dir,"Netlist")
#
if os.path.exists(os.path.join(dir,"Results")):
     print('Results folder already exists and overwriting')
     shutil.rmtree(os.path.join(dir,"Results"))
     os.makedirs('Results',mode=0o777)
New_results_dir =os.path.join(dir,"Results")
# # # for file1 in os.listdir(netlist_dir):
# # #     if file1.endswith('.net'):
# # #         #copyfile(src, dst)
# # #         copyfile(os.path.join(dir,"Netlist",file1),(os.path.join(dir,"Results",file1)))
# #
# # print('copied all .net files')
input_file = os.path.join(dir,"User_input.txt")
# #
with open(input_file) as input:
     input_file = input.readlines()
Location1 = input_file[4].split('=')[1].strip()
Location2 = input_file[5].split('=')[1].strip()
IGBT_Name = input_file[11].split('=')[1].strip()
Diode_Name = input_file[12].split('=')[1].strip()
# #
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
print('Started running simulations')
os.system("sim2 "+os.path.join(netlist_dir, "Script_all_simulations.sxscr"))
print('Fininshed running simulations')
 # print(os.path.join(netlist_dir,"Script_all_simulations.sxscr"))

# x, y = np.loadtxt(os.path.join(New_results_dir, "data_transfer_25C.txt"), delimiter='', unpack=True)
# print('x', list(x)); print('y', list(y))
# plt.plot(x, y, 'ro', label='Transfer characteristics')
# plt.xlabel('VGE(V)')
# plt.ylabel('IC(A)')
# plt.legend()
# plt.show()
for files in os.listdir(New_results_dir):
    if files.endswith('.txt'):
        x, y = [], []
        with open(os.path.join(New_results_dir, files)) as txt_file:
                 raw_data = txt_file.readlines()
                 for itr in range(1, len(raw_data)):
                     x.append(float(raw_data[itr].split()[0].strip()))
                     y.append(float(raw_data[itr].split()[1].strip()))
                 plt.plot(x, y)
                 plt.show()

# transfer_25_file = os.path.join(New_results_dir,"data_transfer_25C.txt")
# x, y = [], []
#
# with open(transfer_25_file) as txtfile1:
#     data1 = txtfile1.readlines()
#     for itr in range(1, len(data1)):
#         # x.append(int(i[0]))
#         # y.append(int(i[1]))
#         x.append(float(data1[itr].split()[0].strip()))
#         y.append(float(data1[itr].split()[1].strip()))
#
# #print(x, y)
#
# plt.plot(x, y)
# #plt.xticks(np.arange(min(str(x)), max(str(x))+1, 1.0))
# #plt.axis([0, 15, 0, 30])
# plt.title('transfer')
# plt.xlabel('VGE(V)')
# plt.ylabel('IC(A)')
# plt.show()

