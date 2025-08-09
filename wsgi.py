import os
from app import app  # Make sure 'app' is imported from your Flask application file

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Default to 5000 if PORT not set
    app.run(host="0.0.0.0", port=port)  # Fixed parameter names (host/port)
