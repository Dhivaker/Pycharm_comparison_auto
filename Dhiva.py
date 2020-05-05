
dir = C:\Users\KumarDhivake\Documents\Students\2019_Dhivaker\Python_automation
netlist_dir = C:\Users\KumarDhivake\Documents\Students\2019_Dhivaker\Python_automation\Netlist

input_file = os.path.join(dir,"User_input.txt")

with open(input_file) as input:
    input_file = input.readlines()

Location1 == input_file[4]
Location2 == input_file[5]
IGBT_Name == input_file[6]
Diode_Name == input_file[7]

for files in os.listdir(netlist_dir):
    if '.net' in files:
        with open(files) as net_file:
            net_file_lines = net_file.readlines()

            for lines in net_file_lines:
                if "<<Location1>>" in lines:
                    lines.replace("<<LocationIGBT>>",Location1)
                if "<<Location2>>" in lines:
                    lines.replace("<<LocationDIODE>>", Location2)
                if "<<IGBT_Name>>" in lines:
                    lines.replace("<<IGBT>>", IGBT_Name)
                if "<<Diode_Name>>" in lines:
                    lines.replace("<<Diode>>", Diode_Name)
    with open("crss_coss.net")as output_file:
        output_file.writelines(net_file_lines)
    with open("design_transfer_27.net")as output_file:
        output_file.writelines(net_file_lines)
    with open("design_transfer_175.net")as output_file:
        output_file.writelines(net_file_lines)
    with open("design_output_27.net")as output_file:
        output_file.writelines(net_file_lines)
    with open("design_output_175.net")as output_file:
            output_file.writelines(net_file_lines)
    with open("diode_vf_27.net")as output_file:
            output_file.writelines(net_file_lines)
    with open("diode_vf_175.net")as output_file:
            output_file.writelines(net_file_lines)

os.system("sim2" +netlist_dir/+".sxscr")






