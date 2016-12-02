'''
Created on 13.10.2016

this is an rpc server to enable the calling if Bloomberg functions remotely

after creating a new remote function you need to write:
     server.register_function(*function_name*)
afterwards to be able to call it remotely

@author: Moustafa Shama
'''

## XML-RPC Service
import sys
import win32serviceutil
import win32service
import win32event
import win32evtlogutil
import win32file
import servicemanager
from xmlrpc.server import SimpleXMLRPCServer

import os
from datetime import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname( __file__ ), '..')))

from rpcs.bloomberg_server  import Bloomberg_Server

    
class XMLRPCSERVICE(win32serviceutil.ServiceFramework):
    _svc_name_ = "BBG_RPCServerService"
    _svc_display_name_ = "BBG_RPCServerService"
    _svc_description_ = "this is the service for rpc server responsible for bbg function calls"
    
    _output_log_file = "c:\service_log_file.txt"
 
    def __init__(self, args):
        win32evtlogutil.AddSourceToRegistry(self._svc_display_name_, sys.executable, "Application")
        win32serviceutil.ServiceFramework.__init__(self, args)
 
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.hSockEvent = win32event.CreateEvent(None, 0, 0, None)
        self.stop_requested = 0

    def SvcStop(self):
        sys.stdout = sys.stderr = open(self._output_log_file, "a")
        print("[RPC Service] [" + datetime.now().strftime("%Y-%m-%d %H:%M:%S") +"] Stopping service")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.stop_requested = 1
        win32event.SetEvent(self.hWaitStop)
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        ## Write a started event
        sys.stdout = sys.stderr = open(self._output_log_file, "a")
        print("[RPC Service] [" + datetime.now().strftime("%Y-%m-%d %H:%M:%S") +"] Starting up")
        
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ' (%s)' % self._svc_name_))
 
        server = SimpleXMLRPCServer(('0.0.0.0', 8080))
        server.register_introspection_functions()
        server.register_multicall_functions()
        
        bbg_srv = Bloomberg_Server(log_file = self._output_log_file)
        
        server.register_instance(bbg_srv)
        
#         self.socket = server.socket
 
        while 1:
            win32file.WSAEventSelect(server, self.hSockEvent,win32file.FD_ACCEPT) 
            rc = win32event.WaitForMultipleObjects((self.hWaitStop,self.hSockEvent), 0, win32event.INFINITE)
            if rc == win32event.WAIT_OBJECT_0:
                break
            else:
                print("[RPC Service] [" + datetime.now().strftime("%Y-%m-%d %H:%M:%S") +"] A request is received")
                server.handle_request()
                win32file.WSAEventSelect(server,self.hSockEvent, 0)
                #server.serve_forever()  ## Works, but breaks the Windows service functionality
 
        ## Write a stopped event
        win32evtlogutil.ReportEvent(self._svc_name_,
                                    servicemanager.PYS_SERVICE_STOPPED,0,
                                    servicemanager.EVENTLOG_INFORMATION_TYPE,
                                    (self._svc_name_,""))

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(XMLRPCSERVICE)