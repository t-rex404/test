#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ======================================================================
class Geometry(object):
    def __init__(
            self,
            text=None,
            width=0,
            height=0,
            left=0,
            top=0):
        """
        Generate a geometry object from the standard Tk geometry string.

        Args:
            text (str): The standard Tk geometry string.
                If str, Must be: `[width]x[height]+[left]+[top]` (integers).
            width (int): The width value.
                If this can be extracted from `text`, this parameter is
                ignored.
            height (int): The height value.
                If this can be extracted from `text`, this parameter is
                ignored.
            left (int): The left value.
                If this can be extracted from `text`, this parameter is
                ignored.
            top (int): The top value.
                If this can be extracted from `text`, this parameter is
                ignored.
        Returns:
            None.

        Examples:
            >>> print(Geometry('1x2+3+4'))
            1x2+3+4
            >>> print(Geometry())
            0x0+0+0
            >>> print(Geometry('1x2'))
            1x2+0+0
            >>> print(Geometry('+1+2'))
            0x0+1+2
            >>> print(Geometry('1+2+3'))
            1x0+2+3
            >>> print(Geometry('1x2+1'))
            1x2+1+0
        """
        self.width, self.height, self.left, self.top = width, height, left, top
        if isinstance(text, str):
            tokens1 = text.split('+')
            tokens2 = tokens1[0].split('x')
            try:
                self.width = int(tokens2[0])
            except (ValueError, IndexError):
                pass
            try:
                self.height = int(tokens2[1])
            except (ValueError, IndexError):
                pass
            try:
                self.left = int(tokens1[1])
            except (ValueError, IndexError):
                pass
            try:
                self.top = int(tokens1[2])
            except (ValueError, IndexError):
                pass

    def __iter__(self):
        k = 'w', 'h', 'l', 't'
        v = self.width, self.height, self.left, self.top
        for k, v, in zip(k, v):
            yield k, v

    def __str__(self):
        return '{w:d}x{h:d}+{l:d}+{t:d}'.format_map(dict(self.items()))

    def __repr__(self):
        return ', '.join([k + '=' + str(v) for k, v in self])

    items = __iter__

    def values(self):
        for k, v in self:
            yield v

    def keys(self):
        for k, v in self:
            yield k

    def set_to_center(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.width // 2 - self.width // 2 + parent.left
        self.top = parent.height // 2 - self.height // 2 + parent.top
        return self

    def set_to_origin(self):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            None.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = 0
        self.top = 0
        return self

    def set_to_top_left(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.left
        self.top = parent.top
        return self

    def set_to_top(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.width // 2 - self.width // 2 + parent.left
        self.top = parent.top
        return self

    def set_to_top_right(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.width - self.width + parent.left
        self.top = parent.top
        return self

    def set_to_right(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.width - self.width + parent.left
        self.top = parent.height // 2 - self.height // 2 + parent.top
        return self

    def set_to_bottom_right(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.width - self.width + parent.left
        self.top = parent.height - self.height + parent.top
        return self

    def set_to_bottom(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.width // 2 - self.width // 2 + parent.left
        self.top = parent.height - self.height + parent.top
        return self

    def set_to_bottom_left(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.left
        self.top = parent.height - self.height + parent.top
        return self

    def set_to_left(self, parent):
        """
        Update the geometry to be centered with respect to a container.

        Args:
            parent (Geometry): The geometry of the container.

        Returns:
            geometry (Geometry): The updated geometry.
        """
        self.left = parent.left
        self.top = parent.height // 2 - self.height // 2 + parent.top
        return self
