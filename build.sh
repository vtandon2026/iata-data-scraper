#!/usr/bin/env bash
# build.sh — runs once during Render deploy to install Chrome + Chromedriver

set -e

echo "==> Installing Chrome and Chromedriver..."

# Install Chrome
wget -q https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
apt-get install -y ./google-chrome-stable_current_amd64.deb
rm google-chrome-stable_current_amd64.deb

# Get matching Chromedriver version
CHROME_VERSION=$(google-chrome --version | grep -oP '\d+\.\d+\.\d+' | head -1)
CHROMEDRIVER_URL="https://storage.googleapis.com/chrome-for-testing-public/${CHROME_VERSION}/linux64/chromedriver-linux64.zip"

echo "==> Chrome: $CHROME_VERSION — downloading Chromedriver..."
wget -q "$CHROMEDRIVER_URL" -O chromedriver.zip || {
    # Fallback: use chromedriver-py to get latest stable
    pip install chromedriver-py --quiet
    CHROMEDRIVER_BIN=$(python -c "import chromedriver_py; print(chromedriver_py.binary_path)")
    echo "CHROMEDRIVER_BIN=$CHROMEDRIVER_BIN" >> /etc/environment
    echo "==> Chromedriver installed via chromedriver-py: $CHROMEDRIVER_BIN"
    exit 0
}

unzip -q chromedriver.zip
mv chromedriver-linux64/chromedriver /usr/local/bin/chromedriver
chmod +x /usr/local/bin/chromedriver
rm -rf chromedriver.zip chromedriver-linux64

echo "==> Chrome + Chromedriver ready"
google-chrome --version
chromedriver --version

# Install Python dependencies
pip install -r requirements.txt