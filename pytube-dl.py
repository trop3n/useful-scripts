from pytube import YouTube

# replace the URL with the desired YouTube video URL
video_url = "https://youtu.be/TnbQNOM7gHg"

try:
    yt = YouTube(video_url)

    # filter for progressive streams (containing both audio and video)
    # and get the highest resolution available within these streams
    stream = (
        yt.streams.filter(progressive=True, file_extension="mp4")
        .order_by("resolution")
        .desc()
        .first()
    )

    if stream:
        print(f"Downloading: {yt.title} ({stream.resolution})")
        stream.download()
        print("Download complete!")
    else:
        print("No suitable streams found.")

except Exception as e:
    print(f"An error occurred: {e}")
