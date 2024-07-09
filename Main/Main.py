import sys
import os

# Add the root directory to sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from GUI.loginGUI import LoginGUI

if __name__ == "__main__":
    app = LoginGUI()
    app.run()
