# This file (node.py) gets copied in several of my projects.  Find out a way 
# to avoid making duplicate copies in each of my projects.

import sys

class NodeType:
    # unknown node type.
    Unknown = 0
    # the document root - typically has only one child element, but it can 
    # have multiple children.
    Root    = 1 
    # node that has name and attributes, and may have child nodes.
    Element = 2
    # node that only has textural content.
    Content = 3

class NodeBase:
    def __init__ (self, nodeType = NodeType.Unknown):
        self.parent = None
        self.__children = []
        self.nodeType = nodeType

    def appendChild (self, node):
        self.__children.append(node)
        node.parent = self

    def appendElement (self, name):
        node = Element(name)
        self.appendChild(node)
        return node

    def appendContent (self, text):
        node = Content(text)
        self.appendChild(node)
        return node

    def firstChild (self):
        return self.__children[0]

    def setChildNodes (self, children):
        self.__children = children

    def getChildNodes (self):
        return self.__children

    def firstChildByName (self, name):
        for child in self.__children:
            if child.nodeType == NodeType.Element and child.name == name:
                return child
        return None

    def getChildByName (self, name):
        children = []
        for child in self.__children:
            if child.nodeType == NodeType.Element and child.name == name:
                children.append(child)
        return children

class Root(NodeBase):
    def __init__ (self):
        NodeBase.__init__(self, NodeType.Root)

class Content(NodeBase):
    def __init__ (self, content):
        NodeBase.__init__(self, NodeType.Content)
        self.content = content

class Element(NodeBase):
    def __init__ (self, name, attrs=None):
        NodeBase.__init__(self, NodeType.Element)
        self.name = name
        self.attrs = attrs
        if self.attrs == None:
            self.attrs = {}

    def getContent (self):
        text = ''
        first = True
        for child in self.getChildNodes():
            if first:
                first = False
            else:
                text += ' '
            if child.nodeType == NodeType.Content:
                text += child.content
            elif child.nodeType == NodeType.Element:
                text += child.getContent()
        return text

    def getAttr (self, name):
        if not self.attrs.has_key(name):
            return None
        return self.attrs[name]

    def setAttr (self, name, val):
        self.attrs[name] = val

    def hasAttr (self, name):
        return self.attrs.has_key(name)

encodeTable = {
    '>': 'gt',
    '<': 'lt',
    '&': 'amp',
    '"': 'quot',
    '\'': 'apos'
}

def encodeString (sin):
    sout = ''
    for c in sin:
        if ord(c) >= 128:
            # encode non-ascii ranges.
            sout += "\\x%2.2x"%ord(c)
        elif encodeTable.has_key(c):
            # encode html symbols.
            sout += '&' + encodeTable[c] + ';'
        else:
            sout += c

    return sout

def convertAttrValue (val):
    if type(val) == type(True):
        if val:
            val = "true"
        else:
            val = "false"
    elif type(val) == type(0):
        val = "%d"%val
    elif type(val) == type(0.0):
        val = "%g"%val

    return val

def prettyPrint (fd, node):
    printNode(fd, node, 0)

def printNode (fd, node, level):
    singleIndent = ' '*4
    indent = singleIndent*level
    if node.nodeType == NodeType.Root:
        # root node itself only contains child nodes.
        for child in node.getChildNodes():
            printNode(fd, child, level)
    elif node.nodeType == NodeType.Element:
        hasChildren = len(node.getChildNodes()) > 0

        # We add '<' and '>' (or '/>') after the element content gets 
        # encoded.
        line = node.name
        if len(node.attrs) > 0:
            keys = node.attrs.keys()
            keys.sort()
            for key in keys:
                val = node.attrs[key]
                if val == None:
                    continue
                val = convertAttrValue(val)
                line += " " + key + '="' + encodeString(val) + '"'

        if hasChildren:
            line = "<%s>\n"%line
            fd.write (indent + line)
            for child in node.getChildNodes():
                printNode(fd, child, level+1)
            line = "</%s>\n"%node.name
            fd.write (indent + line)
        else:
            line = "<%s/>\n"%line
            fd.write (indent + line)

    elif node.nodeType == NodeType.Content:
        content = node.content.strip()
        content = encodeString(content)
        if len(content) > 0:
            fd.write (indent + content + "\n")
