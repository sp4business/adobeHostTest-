import html2text
import requests

def get_url_as_markdown(url):
    response = requests.get(url)
    if response.status_code == 200:
        h = html2text.HTML2Text()
        h.ignore_links = True
        markdown = h.handle(response.text)
        print(markdown)
    else:
        print(f"Failed to retrieve the URL. Status code: {response.status_code}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        get_url_as_markdown(sys.argv[1])
    else:
        print("Please provide a URL as a command-line argument.")
