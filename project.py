import speech_recognition as sr
import pyaudio
import os
import webbrowser as wb
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities , IAudioEndpointVolume

mic = sr.Microphone(device_index=1)
r = sr.Recognizer()
f = sr.Recognizer()

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

d = {i: 0 for i in range(101)}    
d[0] = -65.25
d[1] = -56.992191314697266
d[2] = -51.671180725097656
d[3] = -47.73759078979492
d[4] = -44.61552047729492
d[5] = -42.026729583740234
d[6] = -39.81534194946289
d[7] = -37.88519287109375
d[8] = -36.17274856567383
d[9] = -34.63383865356445
d[10] = -33.23651123046875
d[11] = -31.956890106201172
d[12] = -30.77667808532715
d[13] = -29.681535720825195
d[14] = -28.66002082824707
d[15] = -27.70285415649414
d[16] = -26.80240821838379
d[17] = -25.95233154296875
d[18] = -25.147287368774414
d[19] = -24.38274574279785
d[20] = -23.654823303222656
d[21] = -22.960174560546875
d[22] = -22.295886993408203
d[23] = -21.6594181060791
d[24] = -21.048532485961914
d[25] = -20.461252212524414
d[26] = -19.895822525024414
d[27] = -19.350669860839844
d[28] = -18.824398040771484
d[29] = -18.315736770629883
d[30] = -17.82354736328125
d[31] = -17.3467960357666
d[32] = -16.884546279907227
d[33] = -16.435937881469727
d[34] = -16.000192642211914
d[35] = -15.576590538024902
d[36] = -15.164472579956055
d[37] = -14.763236045837402
d[38] = -14.372318267822266
d[39] = -13.991202354431152
d[40] = -13.61940860748291
d[41] = -13.256492614746094
d[42] = -12.902039527893066
d[43] = -12.55566310882568
d[44] = -12.217005729675293
d[45] = -11.88572883605957
d[46] = -11.561516761779785
d[47] = -11.2440767288208
d[48] = -10.933131217956543
d[49] = -10.62841796875
d[50] = -10.329694747924805
d[51] = -10.036728858947754
d[52] = -9.749302864074707
d[53] = -9.46721076965332
d[54] = -9.190258026123047
d[55] = -8.918261528015137
d[56] = -8.651047706604004
d[57] = -8.388449668884277
d[58] = -8.130311965942383
d[59] = -7.876484394073486
d[60] = -7.626824855804443
d[61] = -7.381200790405273
d[62] = -7.1394829750061035
d[63] = -6.901548862457275
d[64] = -6.6672821044921875
d[65] = -6.436570644378662
d[66] = -6.209307670593262
d[67] = -5.98539400100708
d[68] = -5.764730453491211
d[69] = -5.547224998474121
d[70] = -5.33278751373291
d[71] = -5.121333599090576
d[72] = -4.912779808044434
d[73] = -4.707049369812012
d[74] = -4.5040669441223145
d[75] = -4.3037590980529785
d[76] = -4.1060566902160645
d[77] = -3.9108924865722656
d[78] = -3.718202590942383
d[79] = -3.527923583984375
d[80] = -3.339998245239258
d[81] = -3.1543679237365723
d[82] = -2.970977306365967
d[83] = -2.7897727489471436
d[84] = -2.610703229904175
d[85] = -2.4337174892425537
d[86] = -2.2587697505950928
d[87] = -2.08581280708313
d[88] = -1.9148017168045044
d[89] = -1.7456932067871094
d[90] = -1.5784454345703125
d[91] = -1.4130167961120605
d[92] = -1.2493702173233032
d[93] = -1.0874667167663574
d[94] = -0.9272695183753967
d[95] = -0.768743097782135
d[96] = -0.6118528842926025
d[97] = -0.4565645754337311
d[98] = -0.30284759402275085
d[99] = -0.15066957473754883
d[100] = 0.0

while True:
    with sr.Microphone() as source:
        audio = r.listen(source)
    b = r.recognize_google(audio, language="ru-RU")
    #print(b)
    if b == 'сложение чисел':
        print('скажите, что вам нужно сложить')
        with sr.Microphone() as source:
            Audio = f.listen(source)
        a = f.recognize_google(Audio, language="ru-RU")
        a = a.split()
        summa = int(a[0]) + int(a[2])
        print(summa)
    elif b == "новая вкладка":
        wb.open("http://google.com")
    elif b == "поиск в интернете":
        with sr.Microphone() as source:
            Audio = f.listen(source)
        a = f.recognize_google(Audio, language="ru-RU")
        wb.open("https://google.com/search?q=" + a)
    elif b == "диспетчер задач":
        os.startfile("C:\Windows\System32\Taskmgr.exe")
    elif b == "панель управления":
        os.startfile("C:\Windows\System32\control.exe")
    elif b == "Измени громкость":
        print('назовите желаемую громкость')
        with sr.Microphone() as source:
            Audio = f.listen(source)
        a = f.recognize_google(Audio, language="ru-RU")
        a = int(a)
        if (a>=0 or a<=100):
            if a == 100:
                vol = d[100]
            elif a == 99:
                vol = d[99]
            elif a == 98:
                vol = d[98]
            elif a == 97:
                vol = d[97]
            elif a == 96:
                vol = d[96]
            elif a == 95:
                vol = d[95]
            elif a == 94:
                vol = d[94]
            elif a == 93:
                vol = d[93]
            elif a == 92:
                vol = d[92]
            elif a == 91:
                vol = d[91]
            elif a == 90:
                vol = d[90]
            elif a == 89:
                vol = d[89]
            elif a == 88:
                vol = d[88]
            elif a == 87:
                vol = d[87]
            elif a == 86:
                vol = d[86]
            elif a == 85:
                vol = d[85]
            elif a == 84:
                vol = d[84]
            elif a == 83:
                vol = d[83]
            elif a == 82:
                vol = d[82]
            elif a == 81:
                vol = d[81]
            elif a == 80:
                vol = d[80]
            elif a == 79:
                vol = d[79]
            elif a == 78:
                vol = d[78]
            elif a == 77:
                vol = d[77]
            elif a == 76:
                vol = d[76]
            elif a == 75:
                vol = d[75]
            elif a == 74:
                vol = d[74]
            elif a == 73:
                vol = d[73]
            elif a == 72:
                vol = d[72]
            elif a == 71:
                vol = d[71]
            elif a == 70:
                vol = d[70]
            elif a == 69:
                vol = d[69]
            elif a == 68:
                vol = d[68]
            elif a == 67:
                vol = d[67]
            elif a == 66:
                vol = d[66]
            elif a == 65:
                vol = d[65]
            elif a == 64:
                vol = d[64]
            elif a == 63:
                vol = d[63]
            elif a == 62:
                vol = d[62]
            elif a == 61:
                vol = d[61]
            elif a == 60:
                vol = d[60]
            elif a == 59:
                vol = d[59]
            elif a == 58:
                vol = d[58]
            elif a == 57:
                vol = d[57]
            elif a == 56:
                vol = d[55]
            elif a == 55:
                vol = d[54]
            elif a == 54:
                vol = d[53]
            elif a == 53:
                vol = d[52]
            elif a == 52:
                vol = d[51]
            elif a == 51:
                vol = d[50]
            elif a == 50:
                vol = d[49]
            elif a == 49:
                vol = d[48]
            elif a == 48:
                vol = d[47]
            elif a == 47:
                vol = d[46]
            elif a == 46:
                vol = d[45]
            elif a == 45:
                vol = d[44]
            elif a == 44:
                vol = d[43]
            elif a == 43:
                vol = d[42]
            elif a == 42:
                vol = d[41]
            elif a == 41:
                vol = d[40]
            elif a == 40:
                vol = d[39]
            elif a == 39:
                vol = d[38]
            elif a == 38:
                vol = d[37]
            elif a == 37:
                vol = d[36]
            elif a == 36:
                vol = d[35]
            elif a == 35:
                vol = d[34]
            elif a == 34:
                vol = d[33]
            elif a == 33:
                vol = d[32]
            elif a == 32:
                vol = d[31]
            elif a == 31:
                vol = d[30]
            elif a == 30:
                vol = d[29]
            elif a == 29:
                vol = d[28]
            elif a == 28:
                vol = d[27]
            elif a == 27:
                vol = d[26]
            elif a == 26:
                vol = d[25]
            elif a == 25:
                vol = d[24]
            elif a == 24:
                vol = d[23]
            elif a == 23:
                vol = d[22]
            elif a == 22:
                vol = d[21]
            elif a == 21:
                vol = d[20]
            elif a == 20:
                vol = d[19]
            elif a == 19:
                vol = d[18]
            elif a == 18:
                vol = d[17]
            elif a == 17:
                vol = d[16]
            elif a == 16:
                vol = d[15]
            elif a == 15:
                vol = d[14]
            elif a == 14:
                vol = d[13]
            elif a == 13:
                vol = d[12]
            elif a == 12:
                vol = d[11]
            elif a == 11:
                vol = d[10]
            elif a == 10:
                vol = d[9]
            elif a == 9:
                vol = d[8]
            elif a == 8:
                vol = d[7]
            elif a == 7:
                vol = d[6]
            elif a == 6:
                vol = d[5]
            elif a == 5:
                vol = d[4]
            elif a == 4:
                vol = d[3]
            elif a == 3:
                vol = d[2]
            elif a == 2:
                vol = d[1]
            elif a == 1:
                vol = d[0]
            elif a == 0:
                vol = d[0]
        volume.SetMasterVolumeLevel(vol, None)
        print('установленная громкость = ', a)
    elif b == "выключить голосовой помощник":
        print("Голосовой помощник отключен")
        break



        
