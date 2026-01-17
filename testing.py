from pytube import Search
import webbrowser

def play_first_youtube_video(query):
    s = Search(query)
    video = s.results[0]
    webbrowser.open(video.watch_url)

