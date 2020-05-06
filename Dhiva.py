import os
import shutil

dir = os.getcwd()
netlist_dir =os.path.join(dir,"Netlist")
if os.path.exists(dir):
    print('Results folder already exists')
else:
    os.mkdir('Results')
New_results_dir =os.path.join(dir,"Results")

#print(netlist_dir)
input_file = os.path.join(dir,"User_input.txt")

with open(input_file) as input:
    input_file = input.readlines()

Location1 = input_file[4].split('=')[1].strip()
Location2 = input_file[5].split('=')[1].strip()
IGBT_Name = input_file[11].split('=')[1].strip()
Diode_Name = input_file[12].split('=')[1].strip()

for files in os.listdir(netlist_dir):
    if files.endswith('.net'):
        with open(os.path.join(netlist_dir,files)) as net_file:
            net_file_lines = net_file.readlines()

            for itr in range(0,len(net_file_lines)):
                if "<<LocationIGBT>>" in net_file_lines[itr]:
                    net_file_lines[itr]=net_file_lines[itr].replace("<<LocationIGBT>>", Location1)
                if "<<LocationDIODE>>" in net_file_lines[itr]:
                    net_file_lines[itr]=net_file_lines[itr].replace("<<LocationDIODE>>", Location2)
                if "<<IGBT_Name>>" in net_file_lines[itr]:
                    net_file_lines[itr]=net_file_lines[itr].replace('<<IGBT_Name>>', IGBT_Name)
                if "<<Diode_Name>>" in net_file_lines[itr]:
                    net_file_lines[itr]=net_file_lines[itr].replace("<<Diode_Name>>", Diode_Name)
        with open(os.path.join(netlist_dir, files), 'w')as output_file:
             output_file.writelines(net_file_lines)
    #print(net_file_lines)
os.system("sim2 "+os.path.join(netlist_dir, "Script_all_simulations.sxscr"))
#print(os.path.join(netlist_dir,"Script_all_simulations.sxscr"))





