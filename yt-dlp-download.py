import yt_dlp


def download_video(url, output_path="."):
    options = {
        "outtmpl": f"{output_path}/%(title)s.%(ext)s",
        "format": "bestvideo+bestaudio/best",
        "merge_output_format": "mp4",
        "subtitleslangs": ["en"],
        "writesubtitles": False,
    }

    with yt_dlp.YoutubeDL(options) as ydl:
        ydl.download([url])


# usage
download_video("https://www.youtube.com/watch?v=slPKx67TCG8")
