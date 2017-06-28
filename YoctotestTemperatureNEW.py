import os,sys

import xlwt


from yoctopuce.yocto_api import *
from yoctopuce.yocto_humidity import *
from yoctopuce.yocto_temperature import *
from yoctopuce.yocto_pressure import *

from yoctopuce.yocto_display import *


def LectureTest():
    
    errmsg=YRefParam()
    Meteo = 'METEOMK1-8648D'
    MiniDisplay = 'YD096X16-342A6'
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
        num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='DD.MM.YY HH:MM:SS')
    
    # Setup the API to use local USB devices
    #if YAPI.RegisterHub("127.0.0.1", errmsg)!= YAPI.SUCCESS:
    if YAPI.RegisterHub("usb", errmsg)!= YAPI.SUCCESS:
         sys.exit("init error " + errmsg.value)
         
    if Meteo=='any':
        # retreive any humidity sensor
        sensor = YHumidity.FirstHumidity()
        if sensor is None :
            die('Pas de module connecté')
            m = sensor.get_module()
            Meteo = m.get_serialNumber()
    else:
        m = YModule.FindModule(Meteo)
            
    if not m.isOnline() : die('device not connected')
    humSensor = YHumidity.FindHumidity(Meteo+'.humidity')
    pressSensor = YPressure.FindPressure(Meteo+'.pressure')
    tempSensor = YTemperature.FindTemperature(Meteo+'.temperature')


    if MiniDisplay=='any':
        # retreive any RGB led
        disp = YDisplay.FirstDisplay()
        if disp is None :
            die('No module connected')
    else:
        disp= YDisplay.FindDisplay(MiniDisplay + ".display")
            
    if not disp.isOnline():
         die("Module not connected ")
    # display clean up
    disp.resetAll()
    # retreive the display size
    w=disp.get_displayWidth()
    h=disp.get_displayHeight()
    # retreive the first layer
    l0=disp.get_displayLayer(0)
    l0.clear()
    #display a text in the middle of the screen
    l0.drawText(w / 2,h / 2, YDisplayLayer.ALIGN.CENTER, "Hello world!" )

    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    ligne = 0
    colonne = 0
    from datetime import datetime
    
    while True:
        #ws.write(1, 0, datetime.now(), style1)
        ws.write(ligne, colonne , datetime.now(), style1)
        ws.write(ligne, colonne +1 ,str(tempSensor.get_currentValue()))
        ws.write(ligne, colonne +2 ,str(pressSensor.get_currentValue()))
        ws.write(ligne, colonne +3,str(humSensor.get_currentValue()))
        print('%2.1f' % tempSensor.get_currentValue()+" °C "\
           + "%4.0f" % pressSensor.get_currentValue()+" mb "\
            + "%4.0f" % humSensor.get_currentValue()+"% (Ctrl-c to stop) ")
        l0.clear()
        l0.drawText(w / 2,h / 2, YDisplayLayer.ALIGN.CENTER, str(tempSensor.get_currentValue()) + "°C    " + str(humSensor.get_currentValue())+ "%   "  + str(pressSensor.get_currentValue()) )
        ligne = ligne +1
        wb.save('C:\DATA\JIRA\Python\TEST.xls')
        YAPI.Sleep(2000)
            
def main():
    try:
        LectureTest()
        return 0
    except Exception as err:
        #sys.steder.write('Erreur : %s' % str(err))
        
        return 1
    
if __name__ == '__main__':
    #Strt = main()            
    Srte = LectureTest()
