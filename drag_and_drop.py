
import tkinter as tk
import sys
import win32con
import pythoncom
import pywintypes
import win32com.server.policy
from tkinter.scrolledtext import ScrolledText
from win32com.shell import shell, shellcon



class CustomScrollText(ScrolledText):
    """Custome Scroll Text widget class """

    def __init__(self, x: int, y: int, parent: object = None, **config) -> None:
        super().__init__(parent, **config)
        self.place(x=x, y=y)

    def write(self, value: str) -> None:
        """
        write a given text into output
        """
        self.insert(tk.END, value)
        self.yview_moveto(fraction=1)

    def flush(self, *args) -> None:
        """
        will flush the output
        """


class IDropTarget(win32com.server.policy.DesignatedWrapPolicy):
    """
    The IDropTarget interface is one of the interfaces you implement to provide drag-and-drop
    operations in your application. It contains methods used in any application that can be a
    target for data during a drag-and-drop operation. A drop-target application is responsible for:

        Determining the effect of the drop on the target application.
        Incorporating any valid dropped data when the drop occurs.
        Communicating target feedback to the source so the source application
        can provide appropriate visual feedback such as setting the cursor.
        Implementing drag scrolling.
        Registering and revoking its application windows as drop targets.
        Source:
        https://docs.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-idroptarget

    @Methods:
         Drag Enter
         Drag Over
         Drag Leave
         Drop
    @usage:
         root = tk.Tk()
         ct = ScrollText(parent=root, x=15, y=10)
         you must add write and flush method to  you
         scroll text to redirect the sys.stdout into
         your scroll text widgets
         sys.stdout = ct
         hwnd = root.winfo_id()
         pythoncom.OleInitialize()
         IDropTarget(hwnd)
         root.mainloop()
         """

    _reg_progid_ = "Python.DropTarget"  # register new process
    _reg_clsid_ = "{411c82bc-d2c4-4a67-a8fa-6a94996190bd}"  # process classid
    _reg_desc_ = "OLE DND Drop Target"  # process descriptions
    _com_interfaces_ = (pythoncom.IID_IDropTarget)  # COM Object interface
    _public_methods_ = ('DragEnter', 'DragOver',
                        'DragLeave', 'Drop')  # interface methods
    DATA_FROMAT = (win32con.CF_HDROP, None, pythoncom.DVASPECT_CONTENT, -1,
                   pythoncom.TYMED_HGLOBAL)
    __slots__ = ('hwnd')

    def __init__(self, hwnd: object) -> None:
        """ Initialie the class
        -> args:
            hwnd :Handle to a window(Tkinter root) that can be a
            target for an OLE drag-and-drop operation.
        <-returns:
            None
        """
        self.hwnd = hwnd
        self.drop_effect = shellcon.DROPEFFECT_COPY
        self.data = None
        self._wrap_(self)
        self.register()

    def DragEnter(self, data_object: object, key_state: int,
                  points: tuple, effect: int) -> object:
        """
        Indicates whether a drop can be accepted, and,
        if so, the effect of the drop.
        ->args:
              data_object : A pointer to the IDataObject interface
                 on the data object. This data object contains the
                 data being transferred in the drag-and-drop operation.
                 If the drop occurs, this data object will be incorporated
                 into the target.

              key_state: The current state of the keyboard
                 modifier keys on the keyboard. Possible values
                 can be a combination of any of the flags MK_CONTROL,
                 MK_SHIFT, MK_ALT, MK_BUTTON, MK_LBUTTON, MK_MBUTTON,
                 and MK_RBUTTON.
              points:
                 A POINTL structure containing the current cursor
                 coordinates in screen coordinates.
              effect:
                 On input, pointer to the value of the pdwEffect
                 parameter of the DoDragDrop function. On return,
                 must contain one of the DROPEFFECT flags, which
                 indicates what the result of the drop operation would be.

        """
        try:
            data_object.QueryGetData(self.DATA_FROMAT)
            self.drop_effect = shellcon.DROPEFFECT_COPY
        except pywintypes.com_error:
            self.drop_effect = shellcon.DROPEFFECT_NONE
        return self.drop_effect

    def DragOver(self, key_state: int, point: tuple, effect: int) -> object:
        """
        Provides target feedback to the user and communicates the drop's
        effect to the DoDragDrop function so it can communicate the effect
        of the drop back to the source.
        ->args:
              key_state:
                he current state of the keyboard
                modifier keys on the keyboard. Valid
                values can be a combination of any of the
                flags MK_CONTROL, MK_SHIFT, MK_ALT, MK_BUTTON,
                MK_LBUTTON, MK_MBUTTON, and MK_RBUTTON.
              point:
                A POINTL structure containing the current cursor
                coordinates in screen coordinates.
              effect:
                On input, pointer to the value of the
                pdwEffect parameter of the DoDragDrop function.
                On return, must contain one of the DROPEFFECT flags,
                which indicates what the result of the drop operation would be.
        <- returns:
                object
        """
        return self.drop_effect

    def DragLeave(self) -> None:
        """Removes target feedback and releases the data object."""
        pass

    def Drop(self, data_object: object, key_state: int, points: tuple, effect: int) -> object:
        """
        Incorporates the source data into the target window,
        removes target feedback, and releases the data object.
        ->args:
              data_object:
                A pointer to the IDataObject interface on the data
                object being transferred in the drag-and-drop operation.
              key_state:
                The current state of the keyboard modifier
                keys on the keyboard. Possible values can be a
                combination of any of the flags MK_CONTROL, MK_SHIFT,
                MK_ALT, MK_BUTTON, MK_LBUTTON, MK_MBUTTON, and MK_RBUTTON.
              points:
                A POINTL structure containing the current cursor
                coordinates in screen coordinates.
              effect:
                On input, pointer to the value of the
                pdwEffect parameter of the DoDragDrop function.
                On return, must contain one of the DROPEFFECT flags, .
                which indicates what the result of the drop operation would be.
        <-returns:
                object
        """
        try:
            data_object.QueryGetData(self.DATA_FROMAT)
            self.data = data_object.GetData(self.DATA_FROMAT)
        except pywintypes.com_error as err:
            print(err)
        print(self.get_path())

    def register(self) -> None:
        """
        Registers the specified window as one that can be
        the target of an OLE drag-and-drop operation and
        specifies the IDropTarget instance to use for drop operations.
        """
        try:
            pythoncom.RegisterDragDrop(self.hwnd,
                                       pythoncom.WrapObject(self,
                                                            pythoncom.IID_IDropTarget,
                                                            pythoncom.IID_IDropTarget))
        except pywintypes.com_error as err:
            print(err)

    def get_path(self) -> str:
        """
        returns path of the last draged file
        """
        return ''.join([char for char in shell.DragQueryFileW(
            self.data.data_handle, 0)])


if __name__ == '__main__':
    root = tk.Tk()
    root.title('Test Drag and Drop')
    root.geometry('750x410')
    ct = CustomScrollText(parent=root, x=15, y=10)
    sys.stdout = ct
    hwnd = root.winfo_id()
    pythoncom.OleInitialize()
    IDropTarget(hwnd)
    root.mainloop()

