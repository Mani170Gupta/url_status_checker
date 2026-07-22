from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import os

app = Flask(__name__)

# Restrict this to your actual WordPress domain once live.
# Example: CORS(app, origins=["https://youragencysite.com"])
ALLOWED_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "*")
CORS(app, origins=ALLOWED_ORIGINS if ALLOWED_ORIGINS == "*" else ALLOWED_ORIGINS.split(","))

MAX_URLS_PER_REQUEST = 50
REQUEST_TIMEOUT = 6  # seconds per URL


@app.route("/", methods=["GET"])
def health_check():
    """Simple route to confirm the service is running."""
    return jsonify({"status": "ok", "message": "URL status checker is live"})


@app.route("/check-status", methods=["POST"])
def check_status():
    data = request.get_json(silent=True) or {}
    urls = data.get("urls", [])

    if not isinstance(urls, list) or not urls:
        return jsonify({"error": "Please provide a non-empty list of URLs"}), 400

    urls = urls[:MAX_URLS_PER_REQUEST]

    results = []
    for url in urls:
        url = url.strip()
        if not url:
            continue
        if not url.startswith(("http://", "https://")):
            url = "https://" + url

        try:
            resp = requests.head(
                url,
                allow_redirects=True,
                timeout=REQUEST_TIMEOUT,
                headers={"User-Agent": "Mozilla/5.0 (compatible; StatusChecker/1.0)"},
            )
            if resp.status_code in (405, 403):
                resp = requests.get(
                    url,
                    allow_redirects=True,
                    timeout=REQUEST_TIMEOUT,
                    headers={"User-Agent": "Mozilla/5.0 (compatible; StatusChecker/1.0)"},
                )
            results.append({
                "url": url,
                "status": resp.status_code,
                "final_url": resp.url,
                "redirected": resp.url != url,
            })
        except requests.exceptions.Timeout:
            results.append({"url": url, "status": "Timeout", "error": "Request timed out"})
        except requests.exceptions.ConnectionError:
            results.append({"url": url, "status": "Error", "error": "Connection failed"})
        except Exception as e:
            results.append({"url": url, "status": "Error", "error": str(e)})

    return jsonify({"results": results, "count": len(results)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)