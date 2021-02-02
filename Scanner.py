import speech_recognition as sr
import win32com.client as wincl
import importlib
import serial
import time
import numpy as np
import matplotlib
import os
import sys
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
speak = wincl.Dispatch("SAPI.SpVoice")
arduino = serial.Serial('COM4', 9600)

# Initialize the recognizer  
r = sr.Recognizer()
p=89
flag = 0  

  
while(1):     
      
    # Exception handling to handle 
    # exceptions at the runtime 
    try: 
          
        # use the microphone as source for input. 
        with sr.Microphone() as source2: 
              
            # wait for a second to let the recognizer 
            # adjust the energy threshold based on 
            # the surrounding noise level  
            r.adjust_for_ambient_noise(source2, duration=0.2) 
              
            #listens for the user's input  
            audio2 = r.listen(source2) 
              
            # Using ggogle to recognize audio 
            MyText = r.recognize_google(audio2) 
            MyText = MyText.lower()
            words=MyText.split()
            print(MyText)
            numbers = []
 
           
            for i in words:
                
                 if(i=='radar'):
                      
                    print("scaning initiated", end=" ")
                    speak.Speak('scanning initiated')
                    time.sleep(1)
                    print(".", end=" ")
                    time.sleep(1)
                    print(".", end=" ")
                    time.sleep(1)
                    print(".", end=" ")
                    print(".", end=" ")
                    time.sleep(1)
                    print(".", end=" ")
                    time.sleep(1)
                    print(".", end=" ")
##                    
                    kink = 90
                    pink = 1
                    info = "X{0:.0f}Y{1:.0f}Z".format(kink,pink)
                    arduino.write(info.encode())
                    #print(info)
                    fig = plt.figure(facecolor='k')
                    win = fig.canvas.manager.window # figure window
                    screen_res = win.wm_maxsize() # used for window formatting later
                    dpi = 150.0 # figure resolution
                    fig.set_dpi(dpi) # set figure resolution

                    # polar plot attributes and initial conditions
                    ax = fig.add_subplot(111,polar=True,facecolor='#006d70')
                    ax.set_position([-0.05,-0.05,1.1,1.05])
                    r_max = 100.0 # can change this based on range of sensor
                    ax.set_ylim([0.0,r_max]) # range of distances to show
                    ax.set_xlim([0.0,np.pi]) # limited by the servo span (0-180 deg)
                    ax.tick_params(axis='both',colors='w')
                    ax.grid(color='w',alpha=0.5) # grid color
                    ax.set_rticks(np.linspace(0.0,r_max,5)) # show 5 different distances
                    ax.set_thetagrids(np.linspace(0.0,180.0,10)) # show 10 angles
                    angles = np.arange(0,181,1) # 0 - 180 degrees
                    theta = angles*(np.pi/180.0) # to radians
                    dists = np.ones((len(angles),)) # dummy distances until real data comes in
                    pols, = ax.plot([],linestyle='',marker='o',markerfacecolor = 'w',
                                     markeredgecolor='#EFEFEF',markeredgewidth=1.0,
                                     markersize=10.0,alpha=0.9) # dots for radar points
                    line1, = ax.plot([],color='w',
                                      linewidth=4.0) # sweeping arm plot

                    # figure presentation adjustments
                    fig.set_size_inches(0.96*(screen_res[0]/dpi),0.96*(screen_res[1]/dpi))
                    plot_res = fig.get_window_extent().bounds # window extent for centering
                    win.wm_geometry('+{0:1.0f}+{1:1.0f}'.\
                                    format((screen_res[0]/2.0)-(plot_res[2]/2.0),
                                           (screen_res[1]/2.0)-(plot_res[3]/2.0))) # centering plot
                    fig.canvas.toolbar.pack_forget() # remove toolbar for clean presentation
                    fig.canvas.set_window_title('Arduino Radar')

                    fig.canvas.draw() # draw before loop
                    axbackground = fig.canvas.copy_from_bbox(ax.bbox) # background to keep during loop

                   

                    def stop_event(event):
                        global stop_bool
                        stop_bool = 1
                    prog_stop_ax = fig.add_axes([0.85,0.025,0.125,0.05])
                    pstop = Button(prog_stop_ax,'Stop Program',color='#FCFCFC',hovercolor='w')
                    pstop.on_clicked(stop_event)
                    # button to close window
                    def close_event(event):
                        global stop_bool,close_bool
                        if stop_bool:
                            plt.close('all')
                        stop_bool = 1
                        close_bool = 1
                    close_ax = fig.add_axes([0.025,0.025,0.125,0.05])
                    close_but = Button(close_ax,'Close Plot',color='#FCFCFC',hovercolor='w')
                    close_but.on_clicked(close_event)

                    fig.show()

                 
                    start_word,stop_bool,close_bool = False,False,False
                    while True:
                        try:
                            if stop_bool: # stops program
                                fig.canvas.toolbar.pack_configure() # show toolbar
                                if close_bool: # closes radar window
                                    plt.close('all')
                                break
                            ser_bytes = arduino.readline() # read Arduino serial data
                            decoded_bytes = ser_bytes.decode('utf-8') # decode data to utf-8
                            data = (decoded_bytes.replace('\r','')).replace('\n','')
                            print(data)
                            if start_word:
                                vals = [float(ii) for ii in data.split(',')]
                                if len(vals)<2:
                                    continue
                                angle,dist = vals # separate into angle and distance
                                if dist>r_max:
                                    dist = 0.0 # measuring more than r_max, it's likely inaccurate
                                dists[int(angle)] = dist
                                #print(angle)

                                if angle % 5 ==0: # update every 5 degrees
                                    pols.set_data(theta,dists)
                                    fig.canvas.restore_region(axbackground)
                                    ax.draw_artist(pols)
                                    
                                    line1.set_data(np.repeat((angle*(np.pi/180.0)),2),
                                       np.linspace(0.0,r_max,2))
                                    ax.draw_artist(line1)

                                    fig.canvas.blit(ax.bbox) # replot only data
                                    fig.canvas.flush_events()
                                if angle==p:
                                      
                                      if flag == 0 or 1:
                                          flag = flag + 1
                                      
                                      if flag ==2:
                                          flag ==0
                                          plt.close('all')
                                          print('Scan complete')
                                          speak.Speak('scan complete')
                                          break
                                     # flush for next plot
                            else:
                                if data=='Radar Start': # stard word on Arduno
                                    start_word = True # wait for Arduino to output start word
                                    print("\n")
                                    print('Radar Starting')
                                    speak.Speak('radar shutting')
                                   
                                else:
                                    continue
                               
                        except KeyboardInterrupt:
                            plt.close('all')
                            print("Radar shutting")
                         
                            break

               
                    
                 if(i.isdigit()):
                    numbers.append(int(i))
                    result = int("".join(str(j) for j in numbers))
                    sink = 0
                    data = "X{0:.0f}Y{1:.0f}Z".format(result,sink)
                    
                    #print(data)
                    print("Initiating the turn towards %d degrees" % result)
                    speak.Speak("Initiating the turn towards %d degrees" % result)
                    #checkType('result')
                    arduino.write(data.encode())
                 if(i=='stop'):
                     skunk=90
                     sink = 0
                     data = "X{0:.0f}Y{1:.0f}Z".format(skunk,sink)
                     arduino.write(data.encode())
                     print("Thank you, it was great following ur commands!!")
                     speak.Speak("Thank you, it was great following ur commands!!")
                     sys.exit()
                     
                   
 
              
    except sr.RequestError as e: 
        print("Could not request results; {0}".format(e)) 
          
    except sr.UnknownValueError: 
        print("Waiting for your next command")
