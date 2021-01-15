import win32gui
import win32api
import win32con
import pyautogui
import operator
import random
import time
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
        self.block2pos = dict()
        self.total_pairs = 0

    def init(self):
        """
        Get the handler of the game window and initialize variables
        """
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)

        self.hwnd = hwnd = win32gui.FindWindow(win32con.NULL, self.window_name)
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
        print(self.left_top_and_right_bot)

    def screenshot(self):
        """
        Take the screenshot and cut it into blocks
        """
        image = ImageGrab.grab(self.left_top_and_right_bot)

        image.save(r"C:\Users\dmtalen\Desktop\lianliankan\example3.png")
        # image = Image.open(r"C:\Users\dmtalen\Desktop\lianliankan\example2.png")
        # image.show()

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

    def convert_image_to_matrix(self, debug=False):
        """
        Convert image blocks to according indexes.
        Same image blocks should be assigned to the same indexes.

        init a 2D matrix and a block_index to position dict.
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
                else:
                    image_map.setdefault(this_image_hash, len(image_map) + 1)
                    self.matrix[x][y] = image_map.get(this_image_hash)
                if self.matrix[x][y] != 0:
                    block_idx = self.matrix[x][y]
                    self.block2pos.setdefault(block_idx, [])
                    self.block2pos[block_idx].append((x, y))
        for _, pos_list in self.block2pos.items():
            self.total_pairs += len(pos_list) // 2

        self.print_matrix()

        # for k, v in self.block2pos.items():
        #     print(k, v)

    def find_pairs(self):
        match_pairs = list()
        for block_idx, pos_list in self.block2pos.items():
            length = len(pos_list)
            for i in range(length):
                for j in range(i + 1, length, 1):
                    p1 = pos_list[i]
                    p2 = pos_list[j]
                    if self.match_1(p1, p2):
                        match_pairs.append((p1, p2, 1))
                        # print("1", p1, p2)
                        continue
                    if self.match_2(p1, p2):
                        match_pairs.append((p1, p2, 2))
                        # print("2", p1, p2)
                        continue
                    if self.match_3(p1, p2):
                        match_pairs.append((p1, p2, 3))
                        # print("3", p1, p2)
                        continue
        return match_pairs

    def match_1(self, p1, p2):
        """
        Match method 1.
        The two points are in the same row/column and they can be linked.
        """
        if p1[0] == p2[0]:
            for i in range(min(p1[1], p2[1]) + 1, max(p1[1], p2[1]), 1):
                if self.matrix[p1[0]][i] != 0:
                    return False
            return True
        if p1[1] == p2[1]:
            for i in range(min(p1[0], p2[0]) + 1, max(p1[0], p2[0]), 1):
                if self.matrix[i][p1[1]] != 0:
                    return False
            return True

    def match_2(self, p1, p2):
        """
        Match method 2.
        The two points can be linked with a "7" shape line.
        """
        if p1[0] == p2[0] or p1[1] == p2[1]:
            return False

        """
        cross point: (p2[0], p1[1])
        p2 . . . pc . . . p2
                 .
                 p1
                 .
        p2 . . . pc . . . p2
        
        cross point: (p1[0], p2[1])
        p2       pc       p2
        .                 .
        pc . . . p1 . . . pc
        .                 .        
        p2       pc       p2 
        """
        p_crosses = [(p2[0], p1[1]), (p1[0], p2[1])]
        for p_cross in p_crosses:
            f = [True, True]
            for i in range(min(p1[1], p2[1]) + 1, max(p1[1], p2[1]) + 1, 1):
                if self.matrix[p_cross[0]][i] != 0:
                    f[0] = False
                    break
            for i in range(min(p1[0], p2[0]) + 1, max(p1[0], p2[0]) + 1, 1):
                if self.matrix[i][p_cross[1]] != 0:
                    f[1] = False
                    break
            if f[0] and f[1]:
                return True

    def match_3(self, p1, p2):
        """
        Match method 3.
        The two points can be linked with a "u" shape line.
        """
        def _get_row_empty_points(x, y):
            ret = set()
            i = y - 1
            while i >= 0:
                if self.matrix[x][i] != 0:
                    break
                ret.add((x, i))
                i -= 1
            i = y + 1
            while i < self.col_block_num:
                if self.matrix[x][i] != 0:
                    break
                ret.add((x, i))
                i += 1
            return ret

        def _get_col_empty_points(x, y):
            ret = set()
            i = x - 1
            while i >= 0:
                if self.matrix[i][y] != 0:
                    break
                ret.add((i, y))
                i -= 1
            i = x + 1
            while i < self.row_block_num:
                if self.matrix[i][y] != 0:
                    break
                ret.add((i, y))
                i += 1
            return ret

        p1_row_empty_points = _get_row_empty_points(p1[0], p1[1])
        p2_row_empty_points = _get_row_empty_points(p2[0], p2[1])
        for pa in p1_row_empty_points:
            for pb in p2_row_empty_points:
                f = True
                if pa[1] == pb[1]:
                    for i in range(min(pa[0], pb[0]) + 1, max(pa[0], pb[0]), 1):
                        if self.matrix[i][pa[1]] != 0:
                            f = False
                            break
                    if f:
                        return True

        p1_col_empty_points = _get_col_empty_points(p1[0], p1[1])
        p2_col_empty_points = _get_col_empty_points(p2[0], p2[1])
        for pa in p1_col_empty_points:
            for pb in p2_col_empty_points:
                f = True
                if pa[0] == pb[0]:
                    for i in range(min(pa[1], pb[1]) + 1, max(pa[1], pb[1]), 1):
                        if self.matrix[pa[0]][i] != 0:
                            f = False
                            break
                    if f:
                        return True
        return False

    def execute_one_step(self, p1, p2):
        def _random_shift():
            return random.uniform(0.2, 0.8)

        game_area_left = self.left_top_and_right_bot[0]
        game_area_top = self.left_top_and_right_bot[1]

        from_x = game_area_left + (p1[1] + _random_shift()) * self.block_w
        from_y = game_area_top + (p1[0] + _random_shift()) * self.block_h

        to_x = game_area_left + (p2[1] + _random_shift()) * self.block_w
        to_y = game_area_top + (p2[0] + _random_shift()) * self.block_h

        pyautogui.moveTo(from_x, from_y, 0.3, pyautogui.easeInOutQuad)
        pyautogui.click()

        time.sleep(random.uniform(0.1, 0.8))

        pyautogui.moveTo(to_x, to_y, 0.3, pyautogui.easeInOutQuad)
        pyautogui.click()

    def update_matrix(self, p1, p2):
        self.matrix[p1[0]][p1[1]] = 0
        self.matrix[p2[0]][p2[1]] = 0

    def run(self):
        self.init()
        self.screenshot()
        self.convert_image_to_matrix()

        linked_pairs = 0
        linked_points = set()
        while linked_pairs < self.total_pairs:
            match_pairs = self.find_pairs()
            match_pairs = sorted(match_pairs, key=lambda d: d[2])
            print("!")
            for p1, p2, match_i in match_pairs:
                if p1 in linked_points or p2 in linked_points:
                    continue
                print(p1, p2, match_i)
                self.execute_one_step(p1, p2)
                self.update_matrix(p1, p2)
                linked_pairs += 1
                linked_points.add(p1)
                linked_points.add(p2)
                time.sleep(random.uniform(0.3, 1.2))

    def print_matrix(self):
        for x in range(self.row_block_num):
            tmp = "\t".join([str(i) for i in self.matrix[x]])
            print(tmp)

    @staticmethod
    def color_hash(color):
        """
        Specially for hashing the empty block's color.
        """
        value = ""
        for i in range(5):
            value += "%d,%d,%d," % (color[0], color[1], color[2])
        return hash(value)

    @staticmethod
    def image_hash(img):
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
    robot.run()

