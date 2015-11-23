#!/usr/bin/env python

import sys,wmi,re,multiprocessing as mp
import argparse,functools,socket,csv
from filetimes import filetime_to_dt
DEBUG=False
RPC_SERVER_UNAVAIL=-2147023174
TIMEOUT=3
FIXATTRS="CSName HotFixID InstalledBy InstalledOn Description".split(" ")
def isWindowsListening(computername,debug=False):

    for port in (445,139):
        try:
            s=socket.create_connection((computername,port),timeout=TIMEOUT)
            if debug:
                print "[DD] Can open socket to {} on port {}".format(computername,port)
            s.close()
            return True
        except (socket.error,socket.timeout):
            # if debug:
            pass
        except Exception,e:
            print "[EE] Could not check if {} is listening, because {}".format(computername,e)
    if debug:
        print "[DD] Cannot open any sockets to {}".format(computername)
    return False

def getHotfixes(computername,username=None,password=None,debug=False):
    hotfixes=[]
    # realHostName=None
    try:
        if isWindowsListening(computername,debug) is False:
            return []
        if username and password:            
            w=wmi.WMI(computername,user=username,password=password)
        else:
            w=wmi.WMI(computername)
        
        if debug:
            print "[DD] Connected to {}".format(computername)
        for kb in w.Win32_QuickFixEngineering():
            hfx={}
            # if realHostName is None:
                # realHostName=kb.CSName
            for attr in FIXATTRS:
                hfx[attr]=getattr(kb,attr)
            if len(hfx["InstalledOn"])==16: # hex date
                hfx["InstalledOn"]=filetime_to_dt(int(hfx["InstalledOn"],16)).strftime("%Y/%m/%d")
            if hfx["InstalledOn"].find("/")==-1:
                hfx["InstalledOn"]="{}/{}/{}".format(hfx["InstalledOn"][0:4],hfx["InstalledOn"][4:6],hfx["InstalledOn"][6:8])
            hotfixes.append(hfx)
        if debug:
            print "[DD] Got {} hotfixes from {}".format(len(hotfixes),computername)
        return hotfixes


    except (wmi.x_wmi,socket.timeout,socket.error),x:
        print "[EE] Can't connect to computername: {}".format(computername)
        print x
        return (computername,[])
    except Exception,x:
        print "[EE] Bummer on {}: {}".format(computername,x)
        return (computername,[])




def main():
    global DEBUG
    parser = argparse.ArgumentParser(description='Retrieve HotFixes using WMI based on hostnames')
    parser.add_argument('--processes','-n', type=int,default=5,
                   help='Number of processes to use')
    parser.add_argument('--username','-u', type=str,
                   help='Username')
    parser.add_argument('--password','-p', type=str,
                   help='Password')
    parser.add_argument('computerlist',metavar='FILE',type=str,default='-',nargs='?')
    parser.add_argument('--debug','-d',action='store_true',default=False)
    parser.add_argument('--output','-o', type=str,
                   help='output')  

    args = parser.parse_args()
    DEBUG=args.debug
    if args.debug and args.username:
        print "[DD] using username: {} and pasword {}".format(args.username,args.password)
    computerlist=None
    if args.output is None:
        outfile=sys.stdout
    else:
        outfile=open(args.output,"wb")
    if args.computerlist=='-':
        computerlist=[x.strip() for x in sys.stdin.xreadlines()]
    else:
        with open(args.computerlist,'rb') as f:
            computerlist=[x.strip() for x in f.xreadlines()]
    p=mp.Pool(args.processes)
    f=functools.partial(getHotfixes,debug=args.debug,username=args.username,password=args.password)
    # f("localhost")
    res=p.map(f,computerlist)
    # for e in res:
    #     computername=e[0]
    #     for ip in e[1]:
    #         outfile.write("{}\t{}\n".format(computername,ip))

    csvout=csv.DictWriter(outfile,FIXATTRS,dialect='excel-tab')
    csvout.writeheader()
    # for computer in res:
    for computer in res:
        for row in computer:
            try:
                csvout.writerow(row)
            except Exception,e:
                print "[EE] Bummed out on {}".format(row)




if __name__ == '__main__':
    mp.freeze_support()
    main()
