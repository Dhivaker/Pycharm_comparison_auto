import os
import shutil
import xlrd
import matplotlib.pyplot as plt
import numpy as np
import itertools
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import plotly.graph_objects as go
from reportlab.pdfgen import canvas
from datetime import date
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

class PlotGen:
    #creating structure
    def __init__(self, file_netlist_modifier, plot_editor_excel):
        #creating member variables
        self.dir_cur = os.getcwd()
        self.dir_output_path = os.path.join(self.dir_cur, "output")
        self.dir_netlist = os.path.join(self.dir_cur, "Netlist")
        self.dir_out_netlist = os.path.join(self.dir_output_path, "netlist_files")
        self.dir_out_plot = os.path.join(self.dir_output_path, "output_plots")
        self.file_netlist_modifier = file_netlist_modifier
        self.file_plot_editor = plot_editor_excel
        self.dict_simulated_files_data = {}
        self.dict_table_files_data = {}
        self.dict_plot_labels = {}
        self.dict_table_labels = {}
        self.dict_search_data ={}
        self.dict_data_sheet_value ={}
        self.xlrd_plot_modifier = xlrd.open_workbook(plot_editor_excel)
        #runing initial function
        self.load_plot_labels()
        self.load_table_labels()
        self.load_search_data()
        self.load_data_sheet_value()
        #self.__create_folders()



    def __create_folders(self):
        if os.path.exists(self.dir_output_path):
            print("Previous output exist--> clearing and generating new file")
            shutil.rmtree(self.dir_output_path)
        # create dir for netlistfile
        os.makedirs(self.dir_out_netlist, mode=0o777,exist_ok=True)
        # create dir for plots
        os.makedirs(self.dir_out_plot, mode=0o777,exist_ok=True)

    def load_search_data(self):
        search_data = self.xlrd_plot_modifier.sheet_by_index(2)
        for rows in range(1,search_data.nrows):
            self.dict_search_data[search_data.cell_value(rows,0)] = search_data.cell_value(rows,1)

    def load_data_sheet_value(self):
        data_sheet_value = self.xlrd_plot_modifier.sheet_by_index(2)
        for rows in range(1,data_sheet_value.nrows):
            self.dict_data_sheet_value[data_sheet_value.cell_value(rows,0)] = data_sheet_value.cell_value(rows,2)

    def load_plot_labels(self):
        p_labels_sheet = self.xlrd_plot_modifier.sheet_by_index(1)
        #loading header
        headers = []
        for cell in p_labels_sheet.row(0):
            headers.append(cell.value)
        for rows in range(1,p_labels_sheet.nrows):
            self.dict_plot_labels[p_labels_sheet.cell_value(rows,0)]={}
            for cols in range(1,p_labels_sheet.ncols):
                self.dict_plot_labels[p_labels_sheet.cell_value(rows,0)][headers[cols]]=p_labels_sheet.cell_value(rows,cols)

    def load_table_labels(self):
        t_labels_sheet = self.xlrd_plot_modifier.sheet_by_index(3)
        #loading header
        headers = []
        for cell in t_labels_sheet.row(0):
            headers.append(cell.value)
        for rows in range(1,t_labels_sheet.nrows):
            self.dict_table_labels[t_labels_sheet.cell_value(rows,0)]={}
            for cols in range(1,t_labels_sheet.ncols):
                self.dict_table_labels[t_labels_sheet.cell_value(rows,0)][headers[cols]]=t_labels_sheet.cell_value(rows,cols)

    def __update_file_parameters(self):
        with open(self.file_netlist_modifier) as netlist_file:
            netlist_file_lines = netlist_file.readlines()
        location_1 = netlist_file_lines[4].split('=')[1].strip()
        location_2 = netlist_file_lines[5].split('=')[1].strip()
        igbt_name = netlist_file_lines[11].split('=')[1].strip()
        diode_name = netlist_file_lines[12].split('=')[1].strip()
        for files in os.listdir(self.dir_netlist):
            if files.endswith('.net'):
                with open(os.path.join(self.dir_netlist, files)) as net_file:
                    net_file_lines = net_file.readlines()
                    for itr in range(0, len(net_file_lines)):
                        if "<<LocationIGBT>>" in net_file_lines[itr]:
                            net_file_lines[itr] = net_file_lines[itr].replace("<<LocationIGBT>>", location_1)
                        if "<<LocationDIODE>>" in net_file_lines[itr]:
                            net_file_lines[itr] = net_file_lines[itr].replace("<<LocationDIODE>>", location_2)
                        if "<<IGBT_Name>>" in net_file_lines[itr]:
                            net_file_lines[itr] = net_file_lines[itr].replace('<<IGBT_Name>>', igbt_name)
                        if "<<Diode_Name>>" in net_file_lines[itr]:
                            net_file_lines[itr] = net_file_lines[itr].replace("<<Diode_Name>>", diode_name)
                with open(os.path.join(self.dir_out_netlist, files), 'w') as output_file:
                    output_file.writelines(net_file_lines)

    def generate_net_files(self):
        self.__update_file_parameters()
        print('Started running simulations......Do some meditation')
        os.system("sim2 " + os.path.join(self.dir_netlist, "Script_all_simulations.sxscr"))
        print('Fininshed running simulations')
        # print(os.path.join(netlist_dir,"Script_all_simulations.sxscr"))

    def find_closest_value(self,value,search_list):
        abs_diff = abs(search_list[0]-float(value))
        match_index = 0
        for itr in range(1,len(search_list)):
            if abs_diff > abs(float(value)-search_list[itr]):
                abs_diff = abs(float(value)-search_list[itr])
                match_index = itr
        return match_index


    def find_x1_y1_value(self):
        for plot_type in self.dict_simulated_files_data.keys():
            if plot_type == "output" or plot_type == "diode":
                for files in self.dict_simulated_files_data[plot_type].keys():
                    y1_value = float(self.dict_search_data[files])
                    self.dict_plot_labels[files]["y1value"] = y1_value
                    self.dict_plot_labels[files]["x1value"] = \
                        np.interp(y1_value, self.dict_simulated_files_data[plot_type][files]["y_axis"],
                                  self.dict_simulated_files_data[plot_type][files]["x_axis"])
                    # x1_value = self.dict_plot_labels[files]["x1value"]
                    # self.dict_plot_labels[files]["ty1"] = y1_value + 4
                    # self.dict_plot_labels[files]["tx1"] = x1_value - 5
            if plot_type == "data":
                for files in self.dict_simulated_files_data[plot_type].keys():
                    x1_value = float(self.dict_search_data[files])
                    self.dict_plot_labels[files]["x1value"] = x1_value
                    self.dict_plot_labels[files]["y1value"] = \
                        np.interp(x1_value, self.dict_simulated_files_data[plot_type][files]["x_axis"],
                                  self.dict_simulated_files_data[plot_type][files]["y_axis"])
                    # y1_value = self.dict_plot_labels[files]["y1value"]
                    # self.dict_plot_labels[files]["ty1"] = y1_value + 4
                    # self.dict_plot_labels[files]["tx1"] = x1_value - 5
            if plot_type == "transfer":
                for files in self.dict_simulated_files_data[plot_type].keys():
                    if "transfer_25C.txt" in files:
                        y1_value = float(self.dict_search_data[files])
                        self.dict_plot_labels[files]["y1value"] = y1_value
                        self.dict_plot_labels[files]["x1value"] = \
                            np.interp(y1_value, self.dict_simulated_files_data[plot_type][files]["y_axis"],
                                      self.dict_simulated_files_data[plot_type][files]["x_axis"])
                        x1_value = self.dict_plot_labels[files]["x1value"]
                        self.dict_plot_labels[files]["ty1"] = y1_value+4
                        self.dict_plot_labels[files]["tx1"] = x1_value-5

    def find_x2_y2_value(self):
        for plot_type in self.dict_simulated_files_data.keys():
            if plot_type == "output" or plot_type == "diode":
                for files in self.dict_simulated_files_data[plot_type].keys():
                    x2_value = float(self.dict_data_sheet_value[files])
                    y2_value = float(self.dict_search_data[files])
                    self.dict_plot_labels[files]["y2value"] = y2_value
                    self.dict_plot_labels[files]["x2value"] = x2_value
                    # self.dict_plot_labels[files]["ty1"] = y1_value + 4
                    # self.dict_plot_labels[files]["tx1"] = x1_value - 5
            if plot_type == "data":
                for files in self.dict_simulated_files_data[plot_type].keys():
                    y2_value = float(self.dict_data_sheet_value[files])
                    x2_value = float(self.dict_search_data[files])
                    self.dict_plot_labels[files]["x2value"] = x2_value
                    self.dict_plot_labels[files]["y2value"] = y2_value
                    # self.dict_plot_labels[files]["ty1"] = y1_value + 4
                    # self.dict_plot_labels[files]["tx1"] = x1_value - 5
            if plot_type == "transfer":
                for files in self.dict_simulated_files_data[plot_type].keys():
                    if "transfer_25C.txt" in files:
                        x2_value = float(self.dict_data_sheet_value[files])
                        y2_value = float(self.dict_search_data[files])
                        self.dict_plot_labels[files]["y2value"] = y2_value
                        self.dict_plot_labels[files]["x2value"] = x2_value
                        # self.dict_plot_labels[files]["ty1"] = y1_value+4
                        # self.dict_plot_labels[files]["tx1"] = x1_value-5

    def modify_cies_data(self):
        for files in sorted(os.listdir(self.dir_out_netlist)):
            if "cies.txt" in files:
                self.dict_simulated_files_data["data"]["data_cies.txt"]["x_axis"] = self.dict_simulated_files_data["data"]["data_coss.txt"]["x_axis"]
                y_axis_last_index = len(self.dict_simulated_files_data["data"]["data_cies.txt"]["y_axis"]) - 1
                y_axis_last_value = self.dict_simulated_files_data["data"]["data_cies.txt"]["y_axis"][y_axis_last_index]
                self.dict_simulated_files_data["data"]["data_cies.txt"]["y_axis"] = list(itertools.repeat(y_axis_last_value, len(self.dict_simulated_files_data["data"]["data_cies.txt"]["x_axis"])))

    def load_simulation_data(self):
        for files in sorted(os.listdir(self.dir_out_netlist)):
            if files.endswith('.txt'):
                print("Loading file:", files)
                file_index = files.split('_')[0]
                if file_index not in self.dict_simulated_files_data:
                    self.dict_simulated_files_data[file_index] = {}
                self.dict_simulated_files_data[file_index][files]={}
                self.dict_simulated_files_data[file_index][files]["x_axis"] = []
                self.dict_simulated_files_data[file_index][files]["y_axis"] = []

                with open(os.path.join(self.dir_out_netlist, files)) as txt_file:
                    raw_data = txt_file.readlines()
                    if (len(raw_data[0].split()) > 2):
                        index_x = 1; index_y = 2
                    else:
                        index_x = 0; index_y = 1

                    for itr in range(1, len(raw_data)):
                        self.dict_simulated_files_data[file_index][files]["x_axis"].append(float(raw_data[itr].split()[index_x].strip()))
                        self.dict_simulated_files_data[file_index][files]["y_axis"].append(float(raw_data[itr].split()[index_y].strip()))
        self.find_x1_y1_value()
        self.modify_cies_data()


    def load_table_data(self):
        for images in sorted(os.listdir(self.dir_out_plot)):
            if images.endswith('.png') and not images.startswith('table'):
                print("Loading images:", images)
                self.dict_table_files_data[images] = {}


    def get_plabels(self,filename):
        return self.dict_plot_labels[filename]

    def get_tlabels(self,imagename):
        return self.dict_table_labels[imagename]


    def generate_plot(self):
        for plot_type in self.dict_simulated_files_data.keys():
            for files in self.dict_simulated_files_data[plot_type].keys():
                labels = self.get_plabels(files)
                plt.plot(self.dict_simulated_files_data[plot_type][files]["x_axis"],
                         self.dict_simulated_files_data[plot_type][files]["y_axis"],
                         label=labels["legend"], color=labels["color"], linewidth=labels["width"])
                plt.grid(True)
                plt.ylabel(labels["y_label"])
                plt.xlabel(labels["x_label"])
                plt.yscale(labels["scale"])
                plt.title(labels["title"])
                plt.legend()
                if "output" == plot_type or "diode" == plot_type or "data" == plot_type:
                    plt.axhline(labels["y1value"], color="k", linestyle="--")
                    plt.axvline(labels["x1value"], color="k", linestyle="--")
                    # plt.annotate(labels["text"], xytext=(labels["tx1"], labels["ty1"]),
                    #              xy=(labels["x1value"], labels["y1value"]),
                    #              arrowprops=dict(facecolor='black', shrink=0.05))
                if "transfer_25C" in files:
                    plt.annotate(labels["text"], xytext=(labels["tx1"], labels["ty1"]),
                                 xy=(labels["x1value"], labels["y1value"]),
                                 arrowprops=dict(facecolor='black', shrink=0.05))
                if "data" in files:
                    # plt.yticks(np.arange(10E-14, 10E-9, 10E-1/2))
                    plt.ylim(10E-13, 10E-9)

                # plt.ylim(labels["ymin"], labels["ymax"])

            image_file = plot_type + ".png"
            plt.savefig(os.path.join(self.dir_out_plot, image_file), dpi= 300)
            plt.clf()

    def _generate_table(self):
        for images in self.dict_table_files_data.keys():
                labels = self.get_tlabels(images)
                if "output.png" or "diode.png" in images:
                    fig = go.Figure(data=[go.Table(
                        columnwidth=[250, 150, 350, 350, 200],
                        header=dict(values=[labels["t1st column"], labels["t2nd column"], labels["t3rd column"],
                                            labels["t4th column"], labels["t5th column"]],
                                    line_color='darkslategray',
                                    fill_color='lightskyblue',
                                    align='center',
                                    font=dict(color='black', size=22),
                                    height=60),
                        cells=dict(values=[[labels["1:1 value"], labels["1:2 value"]],
                                           [labels["2:1 value"], labels["2:2 value"]],
                                           [labels["3:1 value"], labels["3:2 value"]],
                                           [labels["4:1 value"], labels["4:2 value"]],
                                           [labels["5:1 value"], labels["5:2 value"]]],
                                   line_color='darkslategray',
                                   fill_color='lightcyan',
                                   align='center',
                                   font=dict(color='black', size=22),
                                   height=40))
                    ])

                    fig.update_layout(width=1000, height=800)
                    table_file = "table_"+images
                    fig.write_image(os.path.join(self.dir_out_plot, table_file))

                if "transfer.png" in images:
                    fig = go.Figure(data=[go.Table(
                        columnwidth=[250, 150, 350, 350, 200],
                        header=dict(values=[labels["t1st column"], labels["t2nd column"], labels["t3rd column"],
                                            labels["t4th column"], labels["t5th column"]],
                                    line_color='darkslategray',
                                    fill_color='lightskyblue',
                                    align='center',
                                    font=dict(color='black', size=22),
                                    height=60),
                        cells=dict(values=[[labels["1:1 value"]],
                                           [labels["2:1 value"]],
                                           [labels["3:1 value"]],
                                           [labels["4:1 value"]],
                                           [labels["5:1 value"]]],
                                   line_color='darkslategray',
                                   fill_color='lightcyan',
                                   align='center',
                                   font=dict(color='black', size=22),
                                   height=40))
                    ])

                    fig.update_layout(width=1000, height=800)
                    fig.write_image(os.path.join(self.dir_out_plot,"table_transfer.png"))

                if "data.png" in images:
                    fig = go.Figure(data=[go.Table(
                        columnwidth=[250, 150, 350, 350, 200],
                        header=dict(values=[labels["t1st column"], labels["t2nd column"], labels["t3rd column"],
                                            labels["t4th column"], labels["t5th column"]],
                                    line_color='darkslategray',
                                    fill_color='lightskyblue',
                                    align='center',
                                    font=dict(color='black', size=22),
                                    height=60),
                        cells=dict(values=[[labels["1:1 value"], labels["1:2 value"], labels["1:3 value"]],
                                           [labels["2:1 value"], labels["2:2 value"], labels["2:3 value"]],
                                           [labels["3:1 value"], labels["3:2 value"], labels["3:3 value"]],
                                           [labels["4:1 value"], labels["4:2 value"], labels["4:3 value"]],
                                           [labels["5:1 value"], labels["5:2 value"], labels["5:3 value"]]],
                                   line_color='darkslategray',
                                   fill_color='lightcyan',
                                   align='center',
                                   font=dict(color='black', size=22),
                                   height=40))
                    ])

                    fig.update_layout(width=1000, height=800)
                    fig.write_image(os.path.join(self.dir_out_plot,"table_data.png"))



    def _generate_pdf_report(self):

        pdf_filename = 'Baredie_report.pdf'
        documentTitle = 'Calibration report'
        pdf = canvas.Canvas(pdf_filename)
        pdf.setTitle(documentTitle)
        logo = 'infineon_logo.png'
        for images in sorted(os.listdir(self.dir_out_plot)):
            if images.endswith('.png') and not images.startswith('infineon') and not images.startswith('table'):
                if "data.png" in images:
                    pdf.drawImage(logo, 420, 800, width=80, height=40, mask='auto')
                    pdf.setFont("Courier-Bold", 28)
                    pdf.drawCentredString(290, 720, "BARE DIE CALIBRATION REPORT")
                    pdf.setFont("Courier-Bold", 20)
                    pdf.drawCentredString(270, 650, "IGBT: IGC11T120X12L_P7351S_2_2")
                    pdf.setFont("Courier-Bold", 20)
                    pdf.drawCentredString(270, 630, "Diode: IDC07D120X8L_L4625C_2_1")
                    pdf.setFont("Courier-Bold", 14)
                    today = date.today()
                    current_date = "Date: %s" % today
                    pdf.drawString(80, 300, current_date)
                    pdf.setFont("Courier-Bold", 14)
                    pdf.drawString(80, 250, "Author: XXX()")
                    pdf.showPage()
                labels = self.get_tlabels(images)
                pdf.drawImage(images, labels["xi_orientation"], labels["yi_orientation"], width=labels["iwidth"],
                              height=labels["iheight"])
                pdf.drawImage(os.path.join(self.dir_out_plot,'table_' + images), labels["xi_orientation"] - 50, labels["yi_orientation"] - 400,
                              width=labels["iwidth"] + 150, height=labels["iheight"] + 150)
                pdf.setFont("Courier-Bold", 24)
                pdf.drawCentredString(300, 700, labels["title"])
                #pdf.setFont("Courier-Bold", 20)
                #pdf.drawCentredString(300, 650, labels["subtitle"])
                pdf.drawImage(logo, 400, 800, width=80, height=40, mask='auto')
                page_num = pdf.getPageNumber()
                text = "Page #%s" % page_num
                pdf.setFont('Times-Roman', 10)
                pdf.drawString(150 * mm, 10 * mm, text)
                print('table_' + images)
                pdf.showPage()

        pdf.save()
