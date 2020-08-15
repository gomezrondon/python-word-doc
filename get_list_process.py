import psutil
# Iterate over all running process


def isWoducmentOpen():
    for proc in psutil.process_iter():
        try:
            # Get process name & pid from process object.
            processName = proc.name()
            processID = proc.pid
            # print(processName , ' ::: ', processID)
            if str(processName) == 'soffice.bin':
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

