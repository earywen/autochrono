import logging
import sys
import os
import webview
import warnings

# Configure logging globally
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Suppress noisy libraries that might cause recursion errors in logs
logging.getLogger('pywebview').setLevel(logging.WARNING)
logging.getLogger('clr').setLevel(logging.WARNING)
logging.getLogger('pythonnet').setLevel(logging.WARNING)

print("--- STARTING OUTLOOK TOOL GEN ---")

# Fix for RecursionError: Disable WebView2 Accessibility features which cause infinite loops
os.environ["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = "--disable-features=Accessibility"
sys.setrecursionlimit(1000)

from vba_generator import UnifiedVBAGenerator

class Api:
    """Python API exposed to JavaScript."""
    
    def __init__(self, window=None):
        self._window = window
        logger.debug("Api initialized")
    
    def set_window(self, window):
        self._window = window
    
    def browse_folder(self):
        """Open folder selection dialog."""
        logger.info("Opening folder dialog...")
        try:
            result = self._window.create_file_dialog(
                webview.FOLDER_DIALOG,
                directory='',
                allow_multiple=False
            )
            if result and len(result) > 0:
                return result[0]
            return None
        except Exception as e:
            logger.error(f"Error in browse_folder: {e}", exc_info=True)
            return None
    
    def browse_file(self):
        """Open file selection dialog for Excel files."""
        logger.info("Opening file dialog...")
        try:
            result = self._window.create_file_dialog(
                webview.OPEN_DIALOG,
                directory='',
                allow_multiple=False,
                file_types=('Excel Files (*.xlsx;*.xls)', 'All Files (*.*)')
            )
            if result and len(result) > 0:
                return result[0]
            return None
        except Exception as e:
            logger.error(f"Error in browse_file: {e}", exc_info=True)
            return None

    def generate_unified_session(self, data):
        """Genere le code unifie pour ThisOutlookSession et le copie dans le presse-papier."""
        try:
            logger.info("Generating Unified Session Code")
            generator = UnifiedVBAGenerator(
                trigram=data.get('trigram', ''),
                chrono_file=data.get('chronoFile', ''),
                chrono_folder=data.get('chronoFolder', '')
            )
            
            code = generator.get_unified_session_module()
            
            # Copie dans le presse papier
            self._copy_to_clipboard(code)
            
            return {'success': True}
            
        except Exception as e:
            logger.error(f"Error generating session: {e}", exc_info=True)
            return {'success': False, 'error': str(e)}
            
    def _copy_to_clipboard(self, text):
        """Copy text to clipboard."""
        logger.debug(f"Copying {len(text)} chars to clipboard")
        try:
            import pyperclip
            pyperclip.copy(text)
        except ImportError:
            import subprocess
            process = subprocess.Popen(['clip'], stdin=subprocess.PIPE)
            process.communicate(text.encode('utf-8'))


def get_html_path():
    """Get the absolute path to the HTML file."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, 'ui', 'index.html')


def main():
    """Launch the OutlookToolGen application."""
    print("--- APPLICATION STARTING ---") 
    warnings.filterwarnings('ignore')
    
    try:
        api = Api()
        
        html_path = get_html_path()
        logger.info(f"Starting application with HTML: {html_path}")
        
        window = webview.create_window(
            title='Outlook Tool Gen (Unified)',
            url=html_path,
            width=950,
            height=700,
            resizable=True,
            js_api=api,
            background_color='#1c365b'
        )
        
        api.set_window(window)
        
        logger.info("Window created, starting webview...")
        webview.start(debug=True, http_server=True)
        
    except Exception as e:
        logger.exception("Critical error during startup")
        input("Press Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"FATAL: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")
