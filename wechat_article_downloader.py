import time
from download_wechat_article import download_wechat_article_from, write_article_to_docx
from get_article_links_from_wechat_history_list_window import get_article_links_from_wechat_history_list_window


def url_is_wechat_article(url):
    return 'mp.weixin.qq.com' in url


def download_one_link(url):
    if url_is_wechat_article(url):
        try:
            article = download_wechat_article_from(url)
            file_name = write_article_to_docx(article, url)
            return True, file_name
        except:
            return False, 'failed'
    else:
        return False, 'not_wechat'


def download_one_link_ui(url):
    print('===============================================\n')
    print('Please wait ...\n')
    success, message = download_one_link(content)
    print('===============================================\n')
    if success:
        print('Success! Saved as {}\n'.format(message))
    else:
        if message == 'failed':
            print('Failed! Please press ENTER to TRY AGAIN.\n')
        elif message == 'not_wechat':
            print('The url is not from wechat article, please download it manually!\n')


def download_from_file_ui(file_name):
    print('===============================================\n')
    links = []
    with open(file_name, 'r') as f:
        for link in f.readlines():
            link = link.strip()
            if link:
                links.append(link)
    wechat_links = []
    other_links = []
    for link in links:
        if url_is_wechat_article(link):
            wechat_links.append(link)
        else:
            other_links.append(link)
    print('Get {} links from "links.txt", '.format(len(links)),
          'where {} are wechat article links, {} are other links.\n'.format(len(wechat_links),
                                                                            len(other_links)))
    success_links = []
    failed_links = []
    print('Start to download wechat links ...\n')
    for i, link in enumerate(wechat_links):
        print('===============================================\n')
        print('Downloading {} / {}, please wait ...\n'.format(i + 1, len(wechat_links)))
        success, message = download_one_link(link)
        if success:
            print('Success {} / {}! Saved as {}\n'.format(i + 1, len(wechat_links), message))
            success_links.append([link, message])
        else:
            print('Failed {} / {}! I will try it later.'.format(i + 1, len(wechat_links)))
            failed_links.append([link, ''])
    print('===============================================\n')
    print('Finished! {} success, {} fails.\n'.format(len(success_links), len(failed_links)))
    if failed_links:
        still_failed_links = []
        print('===============================================\n')
        print('Trying to download failed links again ...\n')
        for i, link in enumerate(failed_links):
            link = link[0]
            print('Downloading {} / {}, please wait ...\n'.format(i + 1, len(failed_links)))
            success, message = download_one_link(link)
            if success:
                print('Success {} / {}! Saved as {}\n'.format(i + 1, len(failed_links), message))
                success_links.append([link, message])
            else:
                print('Still failed {} / {}!\n'.format(i + 1, len(failed_links)))
                still_failed_links.append([link, ''])
        print('===============================================\n')
        print('Finished! Finally {} success, {} fails.\n'.format(len(success_links),
                                                                 len(still_failed_links)))
        failed_links = still_failed_links
    print('===============================================\n')
    print('===============================================')
    print('Success Links')
    print('-----------------------------------------------')
    for link in success_links:
        print(link[0], ', ', link[1])
    print('===============================================')
    print('Failed Links')
    print('-----------------------------------------------')
    for link in failed_links:
        print(link[0], ', ', link[1])
    print('===============================================')
    print('Not wechat Links')
    print('-----------------------------------------------')
    for link in other_links:
        print(link)
    print('-----------------------------------------------\n')


def get_article_links_from_wechat_history_list_window_ui(file_name, append_or_overwrite='overwrite'):
    print('===============================================\n')
    print('Please make sure that the Wechat History List Window is opened, and no other windows covers it!\n')
    print('Please stop using mouse or keyboard until finish!\n')
    print('Will start after 3 secs ...\n')
    time.sleep(3)
    print('===============================================\n')
    print('Start! Please STOP using mouse or keyboard until finish!')
    print('...\n')
    links = get_article_links_from_wechat_history_list_window()
    print('Get {} links from wechat history list window!\n'.format(len(links)))
    print('Writing to {} ...\n'.format(file_name))
    file_mode = 'w' if append_or_overwrite == 'overwrite' else 'a'
    with open(file_name, file_mode) as f:
        for link in links:
            f.write(link + '\n')
    print('Finished! Now you can use your mouse and keyboard happily!\n')


if __name__ == '__main__':

    debug = False
    if debug:
        print('===============================================')
        content = input('Please PASTE the link (url) of a wechat article, or Q to quit.\n')
        article = download_wechat_article_from(content)
        file_name = write_article_to_docx(article, url=content)
        print('Success! Saved as {}!'.format(file_name))
    else:
        while 1:
            print('===============================================\n')
            print('Please PASTE the link (url) of a wechat article,\n',
                  '    or ENTER to quit,\n',
                  '    or R to read links from "links.txt" and save them,\n',
                  '    or C to copy links from wechat history list window to "links.txt" (overwrite),\n',
                  '    or C+ to copy links from wechat history list window to "links.txt" (append),\n',
                  '    or CR to copy & read & save links (equals C then R)')
            content = input('\n')
            if len(content) == 0:
                break
            elif len(content) > 2:
                download_one_link(content)
            elif content.upper() == 'R':
                download_from_file_ui('links.txt')
            elif content.upper() == 'C':
                get_article_links_from_wechat_history_list_window_ui('links.txt', append_or_overwrite='overwrite')
            elif content.upper() == 'C+':
                get_article_links_from_wechat_history_list_window_ui('links.txt', append_or_overwrite='append')
            elif content.upper() == 'CR' or content.upper() == 'RC':
                get_article_links_from_wechat_history_list_window_ui('links.txt', append_or_overwrite='overwrite')
                download_from_file_ui('links.txt')
            elif content.upper() == 'C+R' or content.upper() == 'RC+':
                get_article_links_from_wechat_history_list_window_ui('links.txt', append_or_overwrite='append')
                download_from_file_ui('links.txt')
