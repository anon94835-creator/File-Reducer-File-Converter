import os
from app import app  # Import your Flask app instance

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 4000))  # Default to 4000 if PORT not set
    app.run(host="0.0.0.0", port=port)
