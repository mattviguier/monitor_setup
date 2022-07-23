# %%
import subprocess
import win32api
import win32con
import win32gui
from threading import Thread

# %%
class ProgramWindow:
    def __init__(self, name, run_path, monitor, hwnd=None):
        self.name = name
        self.run_path = run_path
        self.hwnd = hwnd
        self.monitor = monitor
    
    def launch(self):
        self.__run()
        p_thread = Thread(target=self.__wait_active_windows,
                                args=())
        p_thread.start()
        return p_thread
    
    def __run(self):
        subprocess.Popen(self.run_path)

    def __wait_active_windows(self):
        # launch with timeout
        while self.hwnd is None:
            windows_list = get_windows_name_list()
            self.hwnd = get_hwnd_by_name(windows_list, self.name)
    
    def place_window(self, pct_X=0.75, pct_Y=0.75, maximize=True):
        """Position a window on a specific monitor
        
        Args:
            pct_X (float): window width in pourcentage of monitor width 
            pct_Y (float): window height in pourcentage of monitor height 
            maximize (boolean): if true maximize the window
        """
        if self.hwnd is not None:
            monitor_X = self.monitor[0]
            monitor_Y = self.monitor[1]
            monitor_width = self.monitor[2] - monitor_X
            monitor_height = self.monitor[3] - monitor_Y
            
            window_X = int(monitor_X+(1-pct_X)/2*monitor_width)
            window_Y = int(monitor_Y+(1-pct_Y)/2*monitor_height)
            window_width = int(monitor_width*pct_X)
            window_height = int(monitor_height*pct_Y)
            
            # HWND_BOTTOM, HWND_NOTOPMOST, HWND_TOP, HWND_TOPMOST
            # open in center of monitor by default
            win32gui.SetWindowPos(
                self.hwnd,
                win32con.HWND_TOP,
                window_X,
                window_Y,
                window_width,
                window_height,
                win32con.SWP_SHOWWINDOW
            )
            if maximize:
                win32gui.ShowWindow(self.hwnd, win32con.SW_MAXIMIZE)


# %%
def get_windows_name_list():
    """Return list of hwnd visible windows and their name
    """
    windows_list = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            windows_list.append([win32gui.GetWindowText(hwnd), hwnd])

    win32gui.EnumWindows(callback, None)

    return windows_list

# %%
def get_hwnd_by_name(windows_list, name):
    """Get window handle if window text contains name

    Args:
        windows_list (list): list of opened windows
        name (str): name to find in window text
    
    Returns:
        hwnd: window handle
    """
    hwnd = [window for window in windows_list if name.lower() in window[0].lower()]

    if len(hwnd) > 0:
        # only returns the first visible window containing name in their text
        return hwnd[0][1]
    else:
        return None

#%%
def define_monitors(monitors):
    """Define list of monitors with monitor positions and resolutions

    Args:
        monitors (list): list of monitors win32api info
    
    Returns:
        return list of custom defined monitors
    """
    main_monitor, side_monitor, laptop_monitor = None, None, None

    for monitor in monitors:
        monitor_size = win32api.GetMonitorInfo(monitor[0])["Work"]

        if monitor_size[0] == 0 and monitor_size[1] == 0:
            # if top left corner monitor position is (0,0), it's the main monitor
            main_monitor = monitor_size
        elif monitor_size[0] < 0:
            side_monitor = monitor_size
        else:
            laptop_monitor = monitor_size

    return main_monitor, side_monitor, laptop_monitor

# %%
if __name__ == "__main__":
    # get list of monitors
    monitors = win32api.EnumDisplayMonitors()
    main_monitor, side_monitor, laptop_monitor = define_monitors(monitors)
    
    #launch explorer
    explorer = ProgramWindow(name="Explorateur",
                             run_path=["C:/Windows/explorer.exe"],
                             monitor = side_monitor if len(monitors)==3 else main_monitor
                             )

    explorer_thread = explorer.launch()
    explorer_thread.join(timeout=10)
    explorer.place_window()
    
    # outlook = subprocess.Popen(
    #     [
    #         "C:/Program Files (x86)/Microsoft Office/Office16/OUTLOOK.EXE",
    #         "/select",
    #         "outlook:calendar",
    #     ]
    # )


