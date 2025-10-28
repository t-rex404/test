import os
import warnings

import flyingcircus as fc

from pytk import tk
from pytk import msg, dbg
from pytk.Geometry import Geometry


# ======================================================================
def auto_convert(
        text,
        pre_delim=None,
        post_delim=None):
    """
    Convert value to numeric if possible, or strip delimiters from string.

    Args:
        text (str|int|float|complex): The text input string.
        pre_delim (str): initial string decorator.
        post_delim (str): final string decorator.

    Returns:
        val (int|float|complex): The numeric value of the string.

    Examples:
        >>> auto_convert('<100>', '<', '>')
        100
        >>> auto_convert('<100.0>', '<', '>')
        100.0
        >>> auto_convert('100.0+50j')
        (100+50j)
        >>> auto_convert('1e3')
        1000.0
        >>> auto_convert(1000)
        1000
        >>> auto_convert(1000.0)
        1000.0
    """
    if isinstance(text, str):
        if pre_delim and post_delim and \
                fc.has_delim(text, pre_delim, post_delim):
            text = fc.strip_delim(text, pre_delim, post_delim)
        try:
            val = int(text)
        except (TypeError, ValueError):
            try:
                val = float(text)
            except (TypeError, ValueError):
                try:
                    val = complex(text)
                except (TypeError, ValueError):
                    val = text
    else:
        val = text
    return val


# ======================================================================
def get_screen_geometry(from_all=False):
    """
    Workaround to get the size of the current screen in a multi-screen setup.

    Returns:
        geometry (str): The standard Tk geometry string.
            [width]x[height]+[left]+[top]
    """
    root = tk.Tk()
    if from_all:
        width = root.winfo_screenwidth()
        height = root.winfo_screenheight()
        geometry = str(Geometry(width=width, height=height))
    else:
        root.update_idletasks()
        root.attributes('-fullscreen', True)
        root.state('iconic')
        geometry = root.winfo_geometry()
    root.destroy()
    return geometry


# ======================================================================
def set_icon(
        root,
        basename,
        dirpath=os.path.abspath(os.path.dirname(__file__))):
    basepath = os.path.join(dirpath, basename) if dirpath else basename

    # first try modern file formats
    for file_ext in ['png', 'gif']:
        if not basepath.endswith('.' + file_ext):
            filepath = basepath + '.' + file_ext
        else:
            filepath = basepath
        try:
            icon = tk.PhotoImage(file=filepath)
            root.tk.call('wm', 'iconphoto', root._w, icon)
        except tk.TclError:
            warnings.warn('E: Could not use icon `{}`'.format(filepath))
        else:
            return

    # fall back to ico/xbm format
    tk_sys = root.tk.call('tk', 'windowingsystem')
    if tk_sys.startswith('win'):
        filepath = basepath + '.ico'
    else:  # if tk_sys == 'x11':
        filepath = '@' + basepath + '.xbm'
    try:
        root.iconbitmap(filepath)
    except tk.TclError:
        warnings.warn('E: Could not use icon `{}`.'.format(filepath))
    else:
        return


# ======================================================================
def center(target, reference=None):
    target.update_idletasks()
    if reference is None:
        geometry = get_screen_geometry()
    elif not isinstance(reference, (str, Geometry)):
        reference.update_idletasks()
        geometry = reference.winfo_geometry()
    else:
        geometry = reference
    if isinstance(geometry, str):
        geometry = Geometry(geometry)
    target_geometry = Geometry(target.winfo_geometry())
    target.geometry(str(target_geometry.set_to_center(geometry)))


# ======================================================================
def set_aspect(target, parent, aspect=1.0):
    def enforce_aspect_ratio(event):
        width = event.width
        height = int(event.width / aspect)
        if height > event.height:
            height = event.height
            width = int(event.height * aspect)
        target.place(
            in_=parent, x=0, y=0, width=width, height=height)

    parent.bind("<Configure>", enforce_aspect_ratio)
