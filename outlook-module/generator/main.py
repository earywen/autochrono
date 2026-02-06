"""
ChronoCreator Generator - Main entry point with PyWebView
"""

import webview
import os
import sys

sys.setrecursionlimit(500)

from vba_generator import ChronoCreatorGenerator


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
    
    def generate_module(self, data):
        """Generate separate VBA modules based on tool type."""
        try:
            tool_type = data.get('toolType', 'CHRONO')
            
            generator = ChronoCreatorGenerator(
                trigram=data.get('trigram', ''),
                chrono_file=data.get('chronoFile', ''),
                chrono_folder=data.get('chronoFolder', ''),
                user_name=data.get('userName', ''),
                user_phone=data.get('userPhone', '')
            )
            
            if tool_type == 'CHRONO':
                code = generator.get_chrono_module()
                default_name = 'ChronoCreator.bas'
            elif tool_type == 'AR':
                code = generator.get_ar_module()
                default_name = 'AccuseReception.bas'
            else:
                return {'success': False, 'error': 'Type de tool inconnu'}
            
            # Open save dialog
            result = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                directory='',
                save_filename=default_name,
                file_types=('VBA Module (*.bas)', 'All Files (*.*)')
            )
            
            if result:
                # result can be tuple or string depending on platform
                filepath = result[0] if isinstance(result, tuple) else result
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(code)
                return {'success': True, 'path': filepath}
            else:
                return {'success': False, 'error': 'Annule'}
        
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def generate_session(self, data):
        """Generate ThisOutlookSession code."""
        try:
            generator = ChronoCreatorGenerator(
                trigram=data['trigram'],
                chrono_file=data['chronoFile'],
                chrono_folder=data['chronoFolder']
            )
            
            code = generator.get_session_module()
            self._copy_to_clipboard(code)
            return {'success': True, 'type': 'session'}
        
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def _copy_to_clipboard(self, text):
        """Copy text to clipboard."""
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
    """Launch the ChronoCreator Generator application."""
    import warnings
    import logging
    warnings.filterwarnings('ignore')
    logging.getLogger('pywebview').setLevel(logging.CRITICAL)
    
    api = Api()
    
    html_path = get_html_path()
    
    window = webview.create_window(
        title='ChronoCreator Generator',
        url=html_path,
        width=900,
        height=580,
        resizable=True,
        js_api=api,
        background_color='#1c365b'
    )
    
    api.set_window(window)
    
    webview.start(gui='edgechromium', debug=False)


if __name__ == "__main__":
    main()
