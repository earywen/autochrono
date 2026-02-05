"""
AutoChrono - Main entry point with PyWebView
Modern web-based GUI with glassmorphism design
"""

import webview
import os
import sys

# Fix recursion issue with pywebview on Windows
sys.setrecursionlimit(500)

from vba_generator import VBAGenerator


class Api:
    """Python API exposed to JavaScript."""
    
    def __init__(self, window=None):
        self.window = window
    
    def set_window(self, window):
        self.window = window
    
    def browse_folder(self):
        """Open folder selection dialog."""
        result = self.window.create_file_dialog(
            webview.FOLDER_DIALOG,
            directory='',
            allow_multiple=False
        )
        if result and len(result) > 0:
            return result[0]
        return None
    
    def browse_file(self):
        """Open file selection dialog for Excel files."""
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            directory='',
            allow_multiple=False,
            file_types=('Excel Files (*.xlsx;*.xls)', 'All Files (*.*)')
        )
        if result and len(result) > 0:
            return result[0]
        return None
    
    def generate_vba(self, data):
        """Generate VBA code and copy to clipboard."""
        try:
            # Generate VBA code
            generator = VBAGenerator(
                trigram=data['trigram'],
                chrono_folder=data['chronoFolder'],
                chrono_file=data['chronoFile'],
                col_chrono=data['colChrono'],
                col_client=data['colClient'],
                col_trigram=data['colTrigram']
            )
            
            vba_code = generator.get_code()
            
            # Copy to clipboard using pyperclip or fallback
            try:
                import pyperclip
                pyperclip.copy(vba_code)
            except ImportError:
                # Fallback for Windows
                import subprocess
                process = subprocess.Popen(['clip'], stdin=subprocess.PIPE)
                process.communicate(vba_code.encode('utf-8'))
            
            return {'success': True}
        
        except Exception as e:
            return {'success': False, 'error': str(e)}


def get_html_path():
    """Get path to HTML file, works in dev and PyInstaller bundle."""
    if getattr(sys, 'frozen', False):
        # Running as compiled
        base_path = sys._MEIPASS
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, 'ui', 'index.html')


def main():
    """Launch the AutoChrono application."""
    # Suppress pywebview warnings
    import warnings
    import logging
    warnings.filterwarnings('ignore')
    logging.getLogger('pywebview').setLevel(logging.CRITICAL)
    
    api = Api()
    
    html_path = get_html_path()
    
    window = webview.create_window(
        title='AutoChrono',
        url=html_path,
        width=480,
        height=560,
        resizable=False,
        js_api=api,
        background_color='#1c365b'
    )
    
    api.set_window(window)
    
    # Force EdgeChromium backend and disable private mode
    webview.start(gui='edgechromium', debug=False)


if __name__ == "__main__":
    main()
