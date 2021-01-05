import win32gui
from PIL import ImageGrab, Image


class LinkGameRobot:
    def __init__(self):
        self.window_name = "QQ游戏 - 连连看角色版"
        self.top_left_and_bot_right = v = (894, 582, 1482, 965)
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

        img_blocks = \
            [[None] * self.col_block_num for i in range(self.row_block_num)]
        for x in range(1, self.row_block_num):
            for y in range(self.col_block_num):
                top_left_x = x * self.block_w
                top_left_y = y * self.block_h
                bot_right_x = (x + 1) * self.block_w
                bot_right_y = (y + 1) * self.block_h

                box = (top_left_x, top_left_y, bot_right_x, bot_right_y)
                img_block = image.crop(box)
                img_block.show()
            input()



if __name__ == "__main__":
    robot = LinkGameRobot()

    robot.init()
    robot.screenshot()
