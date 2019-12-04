import time
from download_wechat_article import download_wechat_article_from, write_article_to_docx, write_articles_to_one_docx
from get_article_links_from_wechat_history_list_window import get_article_links_from_wechat_history_list_window


def url_is_wechat_article(url):
    return 'mp.weixin.qq.com' in url


def download_one_link(url, save_as_docx=True):
    if url_is_wechat_article(url):
        try:
            article = download_wechat_article_from(url)
            if save_as_docx:
                file_name = write_article_to_docx(article)
                return True, file_name
            else:
                return True, article
        except:
            return False, 'failed'
    else:
        return False, 'not_wechat'


def download_one_link_ui(url):
    print('===============================================\n')
    print('Please wait ...\n')
    success, message = download_one_link(url)
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
    print('            Success Links ', len(success_links))
    print('-----------------------------------------------')
    for link in success_links:
        print(link[0], ', ', link[1])
    print('===============================================')
    print('            Failed Links ', len(failed_links))
    print('-----------------------------------------------')
    for link in failed_links:
        print(link[0], ', ', link[1])
    print('===============================================')
    print('            Not wechat Links ', len(other_links))
    print('-----------------------------------------------')
    for link in other_links:
        print(link)
    print('-----------------------------------------------\n')


def download_from_file_ui_that_links_to_one_docx(file_name):
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
    articles = []
    print('Start to download wechat links ...\n')
    for i, link in enumerate(wechat_links):
        print('===============================================\n')
        print('Downloading {} / {}, please wait ...\n'.format(i + 1, len(wechat_links)))
        success, article = download_one_link(link, save_as_docx=False)
        if success:
            articles.append(article)
            print('Success {} / {}! Downloaded {}\n'.format(i + 1, len(wechat_links), article['title']))
        else:
            articles.append(None)
            print('Failed {} / {}! I will try it later.'.format(i + 1, len(wechat_links)))
    print('===============================================\n')
    failed_num = articles.count(None)
    print('Finished! {} success, {} fails.\n'.format(len(articles) - failed_num, failed_num))
    if None in articles:
        print('===============================================\n')
        print('Trying to download failed links again ...\n')
        failed_i = 0
        for i, article in enumerate(articles):
            if article is None:
                failed_i += 1
                print('Downloading {} / {}, please wait ...\n'.format(failed_i, failed_num))
                success, redownloaded_article = download_one_link(wechat_links[i], save_as_docx=False)
                if success:
                    print('Success {} / {}! Downloaded {}\n'.format(failed_i, failed_num,
                                                                    redownloaded_article['title']))
                    articles[i] = redownloaded_article
                else:
                    print('Still failed {} / {}!\n'.format(failed_i, failed_num))
        print('===============================================\n')
        failed_num = articles.count(None)
        print('Finished! Finally {} success, {} fails.\n'.format(len(articles) - failed_num,
                                                                 failed_num))
    docx_name = write_articles_to_one_docx([a for a in articles if a is not None])
    print('===============================================\n')
    print('===============================================')
    print('            Success Links ', len(articles) - failed_num)
    print('-----------------------------------------------')
    for link, article in zip(wechat_links, articles):
        if article is not None:
            print(link, ', ', article['title'])
    print('===============================================')
    print('            Failed Links ', failed_num)
    print('-----------------------------------------------')
    for link, article in zip(wechat_links, articles):
        if article is None:
            print(link)
    print('===============================================')
    print('            Not wechat Links ', len(other_links))
    print('-----------------------------------------------')
    for link in other_links:
        print(link)
    print('===============================================\n')
    print('Saved as ', docx_name, '\n')
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
        file_name = write_article_to_docx(article)
        print('Success! Saved as {}!'.format(file_name))
    else:
        while 1:
            print('===============================================\n')
            print('Please PASTE the link (url) of a wechat article,\n',
                  '    or ENTER to quit,\n',
                  '    or R to read links from "links.txt" and save them in files (every link relates a file),\n',
                  '    or R1 to read links from "links.txt" and save them in one file (all links in one file),\n',
                  '    or C to copy links from wechat history list window to "links.txt" (overwrite),\n',
                  '    or C+ to copy links from wechat history list window to "links.txt" (append),\n',
                  '    or CR1 equals C then R1, C+R equals C+ then R, etc.')
            content = input('\n').upper()
            link_file_name = 'links.txt'
            if len(content) == 0:
                break
            elif len(content) > 4:
                download_one_link(content)
            else:
                if 'C' in content:
                    if 'C+' in content:
                        get_article_links_from_wechat_history_list_window_ui(link_file_name,
                                                                             append_or_overwrite='append')
                    else:
                        get_article_links_from_wechat_history_list_window_ui(link_file_name,
                                                                             append_or_overwrite='overwrite')
                if 'R' in content:
                    if 'R1' in content:
                        download_from_file_ui_that_links_to_one_docx(link_file_name)
                    else:
                        download_from_file_ui(link_file_name)
