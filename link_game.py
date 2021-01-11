import win32gui
import win32api
import win32con
import operator
from PIL import ImageGrab, Image


class LinkGameRobot:
    def __init__(self):
        self.window_name = "QQ游戏 - 连连看角色版"
        # self.top_left_and_bot_right = v = (894, 582, 1482, 965)
        self.row_block_num = 11
        self.col_block_num = 19

        self.hwnd = None
        self.left_top_and_right_bot = None
        self.block_h = 35.0  # None
        self.block_w = 31.0  # None
        self.img_blocks = \
            [[None] * self.col_block_num for _ in range(self.row_block_num)]
        self.matrix = \
            [[0] * self.col_block_num for _ in range(self.row_block_num)]

    def init(self):
        """
        Get the handler of the game window and initialize variables
        """
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)

        hwnd = win32gui.FindWindow(win32con.NULL, self.window_name)
        if not hwnd:
            exit(-1)

        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)

        window_left, window_top, window_right, window_bottom = \
            win32gui.GetWindowRect(hwnd)
        if min(window_left, window_top) < 0 or \
                window_right > screen_width or \
                window_bottom > screen_height:
            exit(-1)
        window_width = window_right - window_left
        window_height = window_bottom - window_top

        game_area_left = window_left + 14.0 / 800.0 * window_width
        game_area_top = window_top + 181.0 / 600.0 * window_height
        game_area_right = window_left + 603 / 800.0 * window_width
        game_area_bottom = window_top + 566 / 600.0 * window_height

        game_area_width = game_area_right - game_area_left
        game_area_height = game_area_bottom - game_area_top

        self.left_top_and_right_bot = \
            (game_area_left, game_area_top, game_area_right, game_area_bottom)
        self.block_w = game_area_width / self.col_block_num
        self.block_h = game_area_height / self.row_block_num

        print(self.block_w, self.block_h)

    def screenshot(self):
        """
        Take the screenshot and cut it into blocks
        """
        # image = ImageGrab.grab(self.left_top_and_right_bot)
        # image.save(r"C:\Users\dmtalen\Desktop\lianliankan\example2.png")

        image = Image.open(r"C:\Users\dmtalen\Desktop\lianliankan\example2.png")
        image.show()

        for x in range(self.row_block_num):
            for y in range(self.col_block_num):
                top_left_x = y * self.block_w
                top_left_y = x * self.block_h
                bot_right_x = (y + 1) * self.block_w
                bot_right_y = (x + 1) * self.block_h

                box = (top_left_x, top_left_y, bot_right_x, bot_right_y)
                img_block = image.crop(box)
                # img_block.show()
                self.img_blocks[x][y] = img_block

    def convert_image_to_index(self, debug=False):
        """
        Convert image blocks to according indexes.
        Same image blocks should be assigned to the same indexes.
        """
        empty_hash = self.color_hash((48, 76, 112))
        image_map = {}
        # im2index = list()

        for x in range(self.row_block_num):
            for y in range(self.col_block_num):
                # print(x, y)
                this_image = self.img_blocks[x][y]
                this_image_hash = self.image_hash(this_image)
                # this_image.show()
                # print(this_image_hash)
                # input()
                if this_image_hash == empty_hash:
                    self.matrix[x][y] = 0
                    continue
                else:
                    image_map.setdefault(this_image_hash, len(image_map) + 1)
                    self.matrix[x][y] = image_map.get(this_image_hash)

        for x in range(self.row_block_num):
            tmp = " ".join([str(i) for i in self.matrix[x]])
            print(tmp)

    def color_hash(self, color):
        """
        Specially for hashing the empty block's color.
        """
        value = ""
        for i in range(5):
            value += "%d,%d,%d," % (color[0], color[1], color[2])
        return hash(value)

    def image_hash(self, img):
        """
        For the 35 x 31 block, get the diagonal 5, 10, 15, 20, 25 pixels to hash
        """
        value = ""
        for i in range(5, 30, 5):
            c = img.getpixel((i, i))
            value += "%d,%d,%d," % (c[0], c[1], c[2])
        return hash(value)


if __name__ == "__main__":
    robot = LinkGameRobot()

    # robot.init()
    robot.screenshot()
    robot.convert_image_to_index()

