import win32gui
from PIL import ImageGrab, Image


class LinkGameRobot:
    def __init__(self):
        self.window_name = "QQ游戏 - 连连看角色版"
        self.top_left_and_bot_right = v = (893, 581, 1483, 964)
        self.row_block_num = 19
        self.col_block_num = 11
        self.block_h = (v[3] - v[1]) / self.col_block_num  # 34.8
        self.block_w = (v[2] - v[0]) / self.row_block_num  # 31

    def init(self):
        """
        Get the handler of the game window and set it in the foreground
        """
        hwnd = win32gui.FindWindow(0, self.window_name)
        if not hwnd:
            print("Cannot find window %s" % self.window_name)
        win32gui.SetForegroundWindow(hwnd)

    def screenshot(self):
        """
        Take the screenshot and cut it into blocks
        """
        image = ImageGrab.grab(self.top_left_and_bot_right)
        image.show()


if __name__ == "__main__":
    robot = LinkGameRobot()

    robot.init()
    robot.screenshot()
