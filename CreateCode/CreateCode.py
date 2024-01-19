import re


def Step1Data():
    SendValueEventHandler= ""
    data=""
    Step1_data = ""
    CheckSendValueEventHandler = ""
    bool = ""
    Step1_bool = ""
    for i in range(1, num+1):
        SendValueEventHandler = SendValueEventHandler + "static public event SendValueEventHandler SendValueEventHandler" + str(i) + ";" + "\n"
        data = data + "static private string data" + str(i) + ";" + "\n"
        Step1_data = Step1_data + "static public string Step1_data" + str(i) + "\n"\
                                "{" + "\n"\
                                " get" + "\n"\
                                " {" + "\n"\
                                "  return data" + str(i) + ";" + "\n"\
                                " }" + "\n"\
                                " set" + "\n"\
                                " {" + "\n"\
                                "  data" + str(i) + "= value;" + "\n"\
                                "  SendValueEventHandler" + str(i) +"(data" + str(i) + ");" + "\n"\
                                " }" + "\n"\
                                "}" + "\n"
        CheckSendValueEventHandler = CheckSendValueEventHandler + "static public event CheckSendValueEventHandler CheckSendValueEventHandler" + str(i) + ";" + "\n"
        bool = bool + "static private bool bool" + str(i) + ";" + "\n"
        Step1_bool = Step1_bool + "static public bool Step1_bool" + str(i) + "\n"\
                                "{" + "\n"\
                                " get" + "\n"\
                                " {" + "\n"\
                                "  return bool" + str(i) + ";" + "\n"\
                                " }" + "\n"\
                                " set" + "\n"\
                                " {" + "\n"\
                                "  bool" + str(i) + "= value;" + "\n"\
                                "  CheckSendValueEventHandler" + str(i) +"(bool" + str(i) + ");" + "\n"\
                                " }" + "\n"\
                                "}" + "\n"
    return SendValueEventHandler, data, Step1_data, CheckSendValueEventHandler, bool, Step1_bool

def Step1Window():
    Parameter = "public class Parameter" + "\n"\
                "{" + "\n"
    for item in textbox_name:
        Parameter = Parameter + " public string " + item + "_val { get; set; }" + "\n"
    Parameter = Parameter + "}" + "\n"
    LoadConfig = "private void LoadConfig()" + "\n"\
                 "{" + "\n"\
                 " List<Parameter> Parameter_info = Config.Load();" + "\n"
    for item in textbox_name:
        LoadConfig = LoadConfig + " " + item + ".Text = "+"Parameter_info[0]." + item + "_val" + "\n"
    LoadConfig = LoadConfig + "}" + "\n"
    SaveConfig = "private void SaveConfig()" + "\n"\
                 "{" + "\n"\
                 " List<Parameter> Parameter_config = new List<Parameter>()" + "\n"\
                 " {" + "\n"\
                 "  new Parameter() {" + "\n"
    for item in textbox_name:
        SaveConfig = SaveConfig + " " + item + "_val = " + item + ".Text," + "\n"
    SaveConfig = SaveConfig + " }" + "\n"
    SaveConfig = SaveConfig + "};" + "\n"
    SaveConfig = SaveConfig + "Config.Save(Parameter_config);" + "\n"
    SaveConfig = SaveConfig  + "}" + "\n"
    main_content= ""
    for i, item in enumerate(textbox_name):
        main_content = main_content + "private void " + item.split('_Text')[0] + "_Checked(object sender, RoutedEventArgs e)" + "\n"\
                                      "{" + "\n"\
                                      "Step1Data.Step1_bool" + str(i+1) + " = true;" + "\n"\
                                      "}" + "\n"
    return Parameter, LoadConfig, SaveConfig, main_content

def MainWindow():
    CheckSendValueInit = ""
    for i in range(1, num+1):
        CheckSendValueInit = CheckSendValueInit + "Step1Data.CheckSendValueEventHandler" + str(i) + " += (val) =>" + "\n"\
                                                  "{" + "\n"\
                                                  " if (val == true)" + "\n"\
                                                  "{" + "\n"\
                                                  " Step1Data.Step1_data" + str(i) + " = ConvertCoordStr(_downPoint, Display_Image);" + "\n"\
                                                  "}" + "\n"\
                                                  "};" + "\n"
    return CheckSendValueInit
    
with open('Step1Window.txt', 'r', encoding="utf-8") as file:
    lines = file.readlines()

num=0
textbox_name = []
for line in lines:
    match = re.search(r'<TextBox\s.*?x:Name="([^"]+)"', line)
    if match:
        num=num+1
        name_value = match.group(1)
        print(name_value)
        textbox_name.append(match.group(1))


with open('Step1Data.cs'+'.txt', 'w') as file:
    file.write(Step1Data()[0])
    file.write(Step1Data()[1])
    file.write(Step1Data()[2])
    file.write(Step1Data()[3])
    file.write(Step1Data()[4])
    file.write(Step1Data()[5])

with open('Step1Window.xaml.cs'+'.txt', 'w') as file:
    file.write(Step1Window()[0])
    file.write(Step1Window()[1])
    file.write(Step1Window()[2])
    file.write(Step1Window()[3])
        
with open('MainWindow.xaml.cs'+'.txt', 'w') as file:
    file.write(MainWindow())