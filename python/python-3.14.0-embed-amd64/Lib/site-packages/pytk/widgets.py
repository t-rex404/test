import os

# Python interface to Tcl/Tk
from pytk import tk
from pytk import ttk
from pytk import messagebox
from pytk import filedialog
from pytk import simpledialog

from pytk import util

Frame = ttk.Frame
Label = ttk.Label
Button = ttk.Button
Menu = tk.Menu
Progressbar = ttk.Progressbar
Scrollbar = ttk.Scrollbar
Canvas = tk.Canvas
Treeview = ttk.Treeview
Combobox = ttk.Combobox
Entry_ = ttk.Entry
Checkbutton_ = ttk.Checkbutton
# Spinbox = tk.Spinbox
Scale = tk.Scale


# ======================================================================
class Entry(Entry_):
    def __init__(self, *_args, **_kws):
        super(Entry, self).__init__(*_args, **_kws)

    def get_val(self):
        return self.get()

    def set_val(self, val=''):
        try:
            if val is not None:
                val = str(val)
            else:
                raise ValueError
        except ValueError:
            val = ''
        state = self['state']
        self['state'] = 'enabled'
        self.delete(0, tk.END)
        self.insert(0, val)
        self['state'] = state


# ======================================================================
class Checkbutton(Checkbutton_):
    def __init__(self, *_args, **_kws):
        super(Checkbutton, self).__init__(*_args, **_kws)

    def get_val(self):
        return 'selected' in self.state()

    def set_val(self, val=True):
        # if (val and not self.get_val()) or (not val and self.get_val()):
        if bool(val) ^ bool(self.get_val()):  # bitwise xor
            self.toggle()

    def toggle(self):
        self.invoke()


# ======================================================================
class Text(Entry):
    def __init__(self, *_args, **_kws):
        super(Text, self).__init__(*_args, **_kws)

    def get_val(self):
        return self.get()

    def set_val(self, val=''):
        try:
            if val is not None:
                val = str(val)
            else:
                raise ValueError
        except ValueError:
            val = ''
        state = self['state']
        self['state'] = 'enabled'
        self.delete(0, tk.END)
        self.insert(0, val)
        self['state'] = state


# ======================================================================
class Checkbox(Checkbutton):
    def __init__(self, *_args, **_kws):
        super(Checkbox, self).__init__(*_args, **_kws)

    def get_val(self):
        return 'selected' in self.state()

    def set_val(self, val=True):
        # if (val and not self.get_val()) or (not val and self.get_val()):
        if bool(val) ^ bool(self.get_val()):  # bitwise xor
            self.toggle()

    def toggle(self):
        self.invoke()


# ======================================================================
class Spinbox(tk.Spinbox):
    def __init__(self, *_args, **_kws):
        if 'start' in _kws:
            _kws['from_'] = _kws.pop('start')
        if 'stop' in _kws:
            _kws['to'] = _kws.pop('stop')
        if 'step' in _kws:
            _kws['increment'] = _kws.pop('step')
        if 'default' in _kws:
            self.default = _kws.pop('default')
        else:
            self.default = None
        super(Spinbox, self).__init__(*_args, **_kws)
        self.values = _kws['values'] if 'values' in _kws else None
        self.start = _kws['from_'] if 'from_' in _kws else None
        self.stop = _kws['to'] if 'to' in _kws else None
        self.step = _kws['increment'] if 'increment' in _kws else None
        if self.default is not None:
            self.set_val(self.default)
        self.bind('<MouseWheel>', self.mousewheel)
        self.bind('<Button-4>', self.mousewheel)
        self.bind('<Button-5>', self.mousewheel)
        self.sys_events = {
            'scroll_up': {'unix': 4, 'win': +120},
            'scroll_down': {'unix': 5, 'win': -120}}

    def mousewheel(self, event):
        scroll_up = (
                event.num == self.sys_events['scroll_up']['unix'] or
                event.delta == self.sys_events['scroll_up']['win'])
        scroll_down = (
                event.num == self.sys_events['scroll_down']['unix'] or
                event.delta == self.sys_events['scroll_down']['win'])
        if scroll_up:
            self.invoke('buttonup')
        elif scroll_down:
            self.invoke('buttondown')

    def is_valid(self, val=''):
        if self.values:
            result = self.get_val() in self.values
        else:
            result = self.start <= self.get_val() <= self.stop
        return result

    def get_val(self):
        return util.auto_convert(self.get())

    def set_val(self, val=''):
        if self.is_valid(val):
            state = self['state']
            self['state'] = 'normal'
            self.delete(0, tk.END)
            self.insert(0, val)
            self['state'] = state
        else:
            raise ValueError('Spinbox: value `{}` not allowed.'.format(val))


# ======================================================================
class Range(Scale):
    def __init__(self, *_args, **_kws):
        if 'start' in _kws:
            _kws['from_'] = _kws.pop('start')
        if 'stop' in _kws:
            _kws['to'] = _kws.pop('stop')
        if 'step' in _kws:
            _kws['resolution'] = _kws.pop('step')
        if 'default' in _kws:
            self.default = _kws.pop('default')
        else:
            self.default = None
        _kws['showvalue'] = False
        super(Range, self).__init__(*_args, **_kws)
        self.start = _kws['from_'] if 'from_' in _kws else None
        self.stop = _kws['to'] if 'to' in _kws else None
        self.step = _kws['resolution'] if 'resolution' in _kws else None
        if self.default is not None:
            self.set_val(self.default)
        self.bind('<MouseWheel>', self.mousewheel)
        self.bind('<Button-4>', self.mousewheel)
        self.bind('<Button-5>', self.mousewheel)
        self.sys_events = {
            'scroll_up': {'unix': 4, 'win': +120},
            'scroll_down': {'unix': 5, 'win': -120}}

    def mousewheel(self, event):
        scroll_up = (
                event.num == self.sys_events['scroll_up']['unix'] or
                event.delta == self.sys_events['scroll_up']['win'])
        scroll_down = (
                event.num == self.sys_events['scroll_down']['unix'] or
                event.delta == self.sys_events['scroll_down']['win'])
        if scroll_up:
            self.set_val(self.get_val() + self.step)
        elif scroll_down:
            self.set_val(self.get_val() - self.step)

    def is_valid(self, val=0):
        result = self.start <= self.get_val() <= self.stop
        return result

    def get_val(self):
        return util.auto_convert(self.get())

    def set_val(self, val=0):
        if self.is_valid(val):
            state = self['state']
            self.set(val)
            self['state'] = state
        else:
            raise ValueError('Spinbox: value `{}` not allowed.'.format(val))


# ======================================================================
class Listbox(Combobox):
    def __init__(self, *_args, **_kws):
        super(Listbox, self).__init__(*_args, **_kws)
        self['state'] = 'readonly'

    def get_values(self):
        return self.configure('values')[-1]

    def get_val(self):
        return self.get()

    def set_val(self, val=''):
        self.set(val)


# ======================================================================
class Listview(Treeview):
    def __init__(self, *_args, **_kws):
        super(Listview, self).__init__(*_args, **_kws)

    def get_items(self):
        return [self.item(child, 'text') for child in self.get_children('')]

    def add_item(self, item, unique=False):
        items = self.get_items()
        if not unique or unique and item not in items:
            self.insert('', tk.END, text=item)

    def del_item(self, item):
        for child in self.get_children(''):
            if self.item(child, 'text') == item:
                self.delete(child)

    def clear(self):
        for child in self.get_children(''):
            self.delete(child)


# ======================================================================
class ScrollingFrame(Frame):
    def __init__(
            self, parent,
            label_kws=None, label_pack_kws=None,
            *_args, **_kws):
        super(ScrollingFrame, self).__init__(parent, *_args, **_kws)

        if label_kws:
            self.label = Label(parent, **label_kws)
            self.label.pack(**(label_pack_kws if label_pack_kws else {}))

        # create a canvas object and a vertical scrollbar for scrolling it
        self.v_scrollbar = Scrollbar(self, orient='vertical')
        self.v_scrollbar.pack(fill='y', side='right', expand=False)
        self.canvas = Canvas(
            self, bd=0, highlightthickness=0,
            yscrollcommand=self.v_scrollbar.set)
        self.canvas.pack(side='left', fill='both', padx=0, pady=0, expand=True)
        self.v_scrollbar.config(command=self.canvas.yview)

        # reset the view
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

        # create a frame inside the canvas which will be scrolled with it
        self.scrolling = Frame(self.canvas)

        scrolling_id = self.canvas.create_window(
            0, 0, window=self.scrolling, anchor='nw')

        # track changes to the canvas and frame width and sync them,
        # also updating the scrollbar
        def _configure_interior(event):
            # update the scrollbars to match the size of the inner frame
            # size = (
            #     self.scrolling.winfo_reqwidth(),
            #     self.scrolling.winfo_reqheight())
            # canvas.config(scrollregion='0 0 {} {}'.format(*size))
            self.canvas.config(scrollregion=self.canvas.bbox('all'))
            if self.scrolling.winfo_reqwidth() != self.canvas.winfo_width():
                self.canvas.config(width=self.scrolling.winfo_reqwidth())

        self.scrolling.bind('<Configure>', _configure_interior)

        def _configure_canvas(event):
            if self.scrolling.winfo_reqwidth() != self.canvas.winfo_width():
                self.canvas.itemconfigure(scrolling_id,
                                          width=self.canvas.winfo_width())

        self.canvas.bind('<Configure>', _configure_canvas)

        scrolling_widgets = [self.scrolling, self.v_scrollbar]
        for widget in scrolling_widgets:
            widget.bind('<MouseWheel>', self.mousewheel)
            widget.bind('<Button-4>', self.mousewheel)
            widget.bind('<Button-5>', self.mousewheel)
        self.sys_events = {
            'scroll_up': {'unix': 4, 'win': +120},
            'scroll_down': {'unix': 5, 'win': -120}}

    def mousewheel(self, event):
        scroll_up = (
                event.num == self.sys_events['scroll_up']['unix'] or
                event.delta == self.sys_events['scroll_up']['win'])
        scroll_down = (
                event.num == self.sys_events['scroll_down']['unix'] or
                event.delta == self.sys_events['scroll_down']['win'])
        if scroll_up:
            self.canvas.yview_scroll(-1, 'units')
        elif scroll_down:
            self.canvas.yview_scroll(1, 'units')
