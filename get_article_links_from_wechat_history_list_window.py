import time
import pyautogui
import pyperclip
from PIL import Image

ITEM_HEIGHT = 115
SCROLL_NUM = -260
SCROLL_HEIGHT = 120
SEPARATOR_LINE_COLOR = [(237, 237, 237), (242, 242, 242)]
TOP_RIGHT_CORNER_OF_ARTICLE_LIST_IMG = './res/top_right_corner_of_article_list.png'
LINK_COPY_BUTTON_IMG = './res/link_copy_button.png'

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


def get_list_window(click=True):
    screen_img = pyautogui.screenshot()
    top_right_w, top_right_h = get_top_right_corner(screen_img,
                                                    Image.open(TOP_RIGHT_CORNER_OF_ARTICLE_LIST_IMG).convert("RGB"))
    print('Top right corner: {}, {}'.format(top_right_w, top_right_h))
    # Click the top of the wechat article history window to rise it up to the top level
    if click:
        pyautogui.leftClick(x=top_right_w - 90, y=top_right_h + 30, duration=0.5, tween=pyautogui.linear)
    top_left_w, top_left_h = get_not_white(screen_img, top_right_w, top_right_h, -1, 0)
    print('Top left corner: {}, {}'.format(top_left_w, top_left_h))
    middle_right_w, middle_right_h = get_not_white(screen_img, top_right_w, top_right_h, 0, 1)
    middle_right_h += 1
    print('Middle right point: {}, {}'.format(middle_right_w, middle_right_h))
    bottom_right_w, bottom_right_h = get_not_white(screen_img, middle_right_w, middle_right_h, 0, 1)
    print('Bottom right corner: {}, {}'.format(bottom_right_w, bottom_right_h))
    return top_left_w, top_right_w, top_right_h, middle_right_h, bottom_right_h


def get_top_right_corner(screen_img, corner_img):
    left, top, _, _ = pyautogui.locate(needleImage=corner_img, haystackImage=screen_img)
    return left, top


def get_not_white(img, start_w, start_h, direction_w, direction_h):
    img_w, img_h = img.size
    px = img.load()
    w, h = int(start_w), int(start_h)
    w_test, h_test = w + direction_w, h + direction_h
    while 0 <= w_test < img_w and 0 <= h_test < img_h and px[w_test, h_test] == (255, 255, 255):
        w, h = w_test, h_test
        w_test += direction_w
        h_test += direction_h
    return w, h


def scroll_to_top_of_list_window(left, right, middle, bottom):
    samples = []
    list_window = pyautogui.screenshot(region=(left, middle, right - left, bottom - middle))
    px = list_window.load()
    new_samples = [px[int((left + right) / 2), h] for h in range(bottom - middle)]
    while new_samples != samples:
        samples = new_samples
        pyautogui.scroll(-SCROLL_NUM)
        list_window = pyautogui.screenshot(region=(left, middle, right - left, bottom - middle))
        px = list_window.load()
        new_samples = [px[int((left + right) / 2), h] for h in range(bottom - middle)]


def get_separator_lines(img):
    img_w, img_h = img.size
    px = img.load()
    w = img_w // 2
    separator_lines = [h for h in range(img_h) if px[w, h] in SEPARATOR_LINE_COLOR]
    return validate_separator_lines(separator_lines)


def validate_separator_lines(separators):
    validated_separators = []
    for i, s in enumerate(separators):
        if i == 0 or s - validated_separators[-1] == ITEM_HEIGHT:
            validated_separators.append(s)
        else:
            break
    return validated_separators


def cut_list_2_items(img):
    img_w, img_h = img.size
    items = []
    last_item_top = 0
    separator_lines = get_separator_lines(img)
    # print('Separator: ', separator_lines)
    for sl in separator_lines[0: -1]:
        items.append(img.crop((0, sl, img_w, sl + ITEM_HEIGHT - 2)))
        last_item_top = sl
    sl = separator_lines[-1]
    if sl + ITEM_HEIGHT == img_h + 1:
        items.append(img.crop((0, sl, img_w, sl + ITEM_HEIGHT - 2)))
        last_item_top = sl
    return items, last_item_top


def get_items_from_list_window(left, right, middle, bottom, scroll_times):
    list_windows = [pyautogui.screenshot(region=(left, middle, right - left, bottom - middle))]
    items, last_item_top = cut_list_2_items(list_windows[-1])
    action_history = ['item'] * len(items)
    actual_scroll_times = 0
    actual_scroll_distance, expected_scroll_distance = 0, 0
    while actual_scroll_distance == expected_scroll_distance:
        for i in range(scroll_times):
            pyautogui.scroll(SCROLL_NUM, pause=0.3)
            actual_scroll_times += 1
            action_history.append('scroll')
        time.sleep(1.5)

        list_windows.append(pyautogui.screenshot(region=(left, middle, right - left, bottom - middle)))
        anchor = pyautogui.locate(needleImage=items[-1], haystackImage=list_windows[-1])
        if anchor is None:
            items[-1].show()
            list_windows[-1].show()
        new_items, new_last_item_top = cut_list_2_items(list_windows[-1].crop((0, anchor.top + ITEM_HEIGHT,
                                                                               right - left, bottom - middle)))
        new_last_item_top += anchor.top + ITEM_HEIGHT
        actual_scroll_distance = last_item_top - anchor.top
        expected_scroll_distance = scroll_times * SCROLL_HEIGHT

        items += new_items
        action_history += ['item'] * len(new_items)
        last_item_top = new_last_item_top

        print("actual_scroll_distance: {}, expected_scroll_distance: {}".format(actual_scroll_distance,
                                                                                expected_scroll_distance))

    for i in range(actual_scroll_times):
        pyautogui.scroll(-SCROLL_NUM, pause=0.3)

    return items, action_history


def open_article_and_copy_link(left, right, middle, bottom, items, action_history, mouse_w_in_list_window,
                               mouse_h_in_list_window):
    article_links = []
    have_scrolled = True
    list_window = None
    for action in action_history:
        if action == 'scroll':
            pyautogui.moveTo(x=mouse_w_in_list_window, y=mouse_h_in_list_window, tween=pyautogui.linear)
            pyautogui.scroll(SCROLL_NUM, pause=0.3)
            have_scrolled = True
        elif action == 'item':
            item = items.pop(0)
            if have_scrolled:
                list_window = pyautogui.screenshot(region=(left, middle, right - left, bottom - middle))
                have_scrolled = False
            relative_w, relative_h = pyautogui.center(pyautogui.locate(needleImage=item, haystackImage=list_window))
            item_w = relative_w + left
            item_h = relative_h + middle
            pyautogui.leftClick(x=item_w, y=item_h, tween=pyautogui.linear)
            link_copy_button_w, link_copy_button_h = pyautogui.center(pyautogui.locateOnScreen(LINK_COPY_BUTTON_IMG))
            pyautogui.leftClick(x=link_copy_button_w, y=link_copy_button_h)
            article_links.append(pyperclip.paste())
            pyautogui.keyDown('altleft', pause=0.3)
            pyautogui.press('f4')
            pyautogui.keyUp('altleft', pause=0.3)
    return article_links


def get_article_links_from_wechat_history_list_window():
    left, right, top, middle, bottom = get_list_window()
    time.sleep(1)
    mouse_w_in_scroll_area = (left + right) // 2
    mouse_h_in_scroll_area = middle + 50
    pyautogui.moveTo(x=mouse_w_in_scroll_area, y=mouse_h_in_scroll_area, duration=0.5, tween=pyautogui.linear)
    scroll_to_top_of_list_window(left, right, middle, bottom)

    scroll_times = (bottom - middle) // SCROLL_HEIGHT - 2
    items, action_history = get_items_from_list_window(left, right, middle, bottom, scroll_times)
    time.sleep(1)

    article_links = open_article_and_copy_link(left, right, middle, bottom, items, action_history,
                                               mouse_w_in_scroll_area, mouse_h_in_scroll_area)
    return article_links


if __name__ == '__main__':
    article_links = get_article_links_from_wechat_history_list_window()
    for link in article_links:
        print(link)
