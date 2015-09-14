# -*- coding: utf8 -*-

#################################################################
# post_extractor was written in its entirety by Dan Stuart      #
# And is distributed under Creative Commons by-nc-sa            #
# (Attribution, non-commercial, share-alike).                   #
# Contact me at dstu@umich.edu or drestuart@gmail.com           #
# With any questions.  5/17/2010                                #
###                                                             #
### @   A post extractor for discuss.hk base on post_extractor  #
###         which created by Dan Stuart                         #
### @   Download openpyxl from write and read execl for python  #
###     command-> python -m pip install openpyxl                #
### @   Run this program                                        #
###     Command-> python post_extractor.py                      #
#################################################################

from re import search, findall, sub, S
import sys


def parse_post(html):
    '''Extracts the usernames and posts from the input html document.'''
    
    #Eliminate quote blocks
    html = sub("<!--QuoteBegin-->[\w\W]+?<!--QuoteEBegin-->", '', html)
    html = sub("<!--QuoteEnd-->[\w\W]+?<!--QuoteEEnd-->", '', html)

    #Find usernames
    title_html = ur"<\/a>(.+?)\s+</h1>\s+<!--overture start-->"
    title = findall(title_html, html)

    #Find usernames
    username_html = ur"<a href=\"space.php\?uid=\d+\" target=\"_blank\">(.+?)</a>"
    usernames = findall(username_html, html)


    #Find time
    #time_html = ur".+t_smallfont.+<\/em>\s+.+?\s(.+?)&nbsp;\s+<a href=\"viewthread\.php\?tid"
    #hour_html = ur".+t_smallfont.+<\/em>\s+.+?\s.+?\s(.+?)&nbsp;\s+<a href=\"viewthread\.php\?tid"
    #date_html = ur".+t_smallfont.+<\/em>\s+.+?\s(.+?)\s.+&nbsp;\s+<a href=\"viewthread\.php\?tid"
    time_html = ur"(\d+-\d+-\d+\s\d\d:\d\d\s[A|P]M)&nbsp;"
    hour_html = ur"\d+-\d+-\d+\s(\d\d:\d\d\s[A|P]M)&nbsp;"
    date_html = ur"(\d+-\d+-\d+)\s\d\d:\d\d\s[A|P]M&nbsp;"
    times = findall(time_html, html)
    dates = findall(date_html, html)
    hours_12 = findall(hour_html, html)
    hours_24 = []
    for hour in hours_12:
        tmp = search(r'\d+', hour)
        if search(r'PM', hour):
            if int(tmp.group()) == 12:
                hours_24.append(int(tmp.group()) + 0)
            else:         
                hours_24.append(int(tmp.group()) + 12)
        elif search(r'AM', hour):
            if int(tmp.group()) == 12:
                hours_24.append(int(tmp.group()) - 12)
            else:         
                hours_24.append(int(tmp.group()) + 0)


    #Find post id
    id_html = ur"\'\)\">(.+)<sup>#<\/sup>"
    postids = findall(id_html, html)


    #Find post
    post_html = ur"<div id=\"postmessage_.+?>(.+?)</div><!--msg\?-->|<div class=\"notice\" .+?>(.+?)</div>"
    posts = findall(post_html, html, S)
    
    retposts = []
    quotes = []
    for post in posts:
        tmp_post = []
        for string in post:
            string = sub("<br />|<br/>|<br>", '', string)
            string = sub("</span><b>.+?</font></a>", '', string)
            string = sub("<a href=\"http://www.discuss.com.hk/iphone.+?</a>", '', string)
            string = sub("<a href=\"http://www.discuss.com.hk/android\".+?</a>", '', string)
            string = sub("<a .+?mobile.jpg.+?</a>", '', string)
            string = sub("<img .+?iphoneD.jpg.+?/>", '', string)
            string = sub("<img .+?androidD.jpg.+?/>", '', string)
            string = sub("<img .+?mobile.jpg.+?/>", '', string)
            string = sub("<a .+?back.gif.+?</a>", '', string)
            string = sub("<a.+?yahoo.+?</a>", '', string)
            string = sub("<!-- Ad space:.+?addStaticSlot\"></script>", '', string)
            
            string = sub("<img.+?smilieid.+?>", ' ## ICON ## ', string)           # change icon
            string = sub("(<span id=\"postorig_.+?>)", r'\1\n', string)     # line break after <span>
            string = sub("(<div class=\"quote\">)", r'\n\1', string)        # line break after <div>

            if string:
                tmp_quote = search('<div.+<blockquote>((?s).*)<\/div>', string)
                if search(r'quote', string):
                    try:
                        quote_text = tmp_quote.group()                     
                        quote_text = sub("\r?\n|\r", ' ', quote_text)
                        quotes.append(quote_text)
                    except:
                        quotes.append("Fetch Quote Problem")
                else:
                    quotes.append("")

            string = sub("<div.+<blockquote>(?s).*</div>", ' ## QUOTE ## ', string)  # del quote text  
            string = sub("<.+?>", '', string)
            string = sub("\r?\n|\r", ' ', string)
            
            

            tmp_post.append(string)

        retposts.append(tmp_post)

    return usernames, retposts, times, postids, title, dates, hours_24, quotes
