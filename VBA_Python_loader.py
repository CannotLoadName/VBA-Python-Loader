#-*-coding:utf-8;-*-
if __name__=="__main__":
    from errno import ESRCH
    from pythoncom import GetRunningObjectTable,IID_IDispatch
    from sys import argv
    from win32com.client import Dispatch
    scriptsource=argv[1]
    monikerjudge=argv[2]
    programid=argv[3]
    rot=GetRunningObjectTable()
    monikers=rot.EnumRunning()
    iterating=True
    while iterating:
        monikertuple=monikers.Next()
        if monikertuple:
            for moniker in monikertuple:
                try:
                    dispatch=Dispatch(rot.GetObject(moniker).QueryInterface(IID_IDispatch),programid)
                except:
                    pass
                else:
                    namespace={"app":dispatch,"application":dispatch}
                    try:
                        judgeresult=eval(monikerjudge,namespace)
                    except:
                        pass
                    else:
                        if judgeresult:
                            if scriptsource.startswith(("-i","--ipython")):
                                from IPython import start_ipython
                                start_ipython(("-i","--pylab=tk"),user_ns=namespace)
                            else:
                                exec(eval(scriptsource,namespace),namespace)
                            iterating=False
                            break
        else:
            raise ProcessLookupError(ESRCH,"The COM object was not found.")