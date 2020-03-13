plpump -f:photolistfile -t:transformfile -o:outputdirectory [-u:username -p:password] [-v] [-d]

Photo list pump selects photos from a series of inputs specified in a photo list file (photolistfile) applying a transformation (transformfile) along the ways and copy them into an output directory (outputdirectory).  The inputs in the photo list files are selected according to a series of complex rules, including normalization of photos across several sources.

-f:photolistfile = XML file containing rules to select photos (e.g. c:\photolist.xml).  Example below.
-t:transformfile = XML file containing transformation rules (e.g. c:\transform.xml).  Example below.
-o:outputdirectory = Directory to copy photos into (e.g. c:\output\). 
-u:username = Username for .net Passport to select photos from an MSN Group (only necessary if an MSN Group is specified in the photolist file) (e.g. arnold_blinn@hotmail.com)
-p:password = Password for .net Passport to select photos from an MSN Group (only necessary if an MSN Group is specified in the photolist file) (e.g. xxxxx)
-v = Verbose mode (default is false)
-d = Delete files in output directory (default is false)

Transform File Example:

<Format>
    <DesiredWidth>1024</DesiredWidth>
    <DesiredHeight>768</DesiredHeight>
    <TopMargin>10</TopMargin>
    <LeftMargin>10</LeftMargin>
    <BottomMargin>10</BottomMargin>
    <RightMargin>10</RightMargin>
    <ColorMargin>White</ColorMargin>
    <HAlign>1</HAlign>
    <VAlign>1</VAlign>
    <Grow>1</Grow>
    <Shrink>1</Shrink>
    <Rotate>0</Rotate>
    <Pad>0</Pad>
    <ColorPad>black</ColorPad>
    <Quality>100</Quality>
    <Sharpen>1</Sharpen>
</Format>

Where:

DesiredWidth = Width of final image
DesiredHeight = Height of final image
TopMargin, LeftMargin, BottomMargin, RightMargin = Width, in pixels, of a margin around the photo
ColorMargin = Color of the margin 
Grow = Grow the photo to the desired size
Shrink = Shrink the photo to the desired size
Pad = Pad to the desired width or height
HAlign = Horizontal Alighment of photo in desired width and height if the photo is grown or shrunk and padded
VAlign = Vertical Alignment
Rotate = Rotate to optimally fill the desired width and height
ColorPad = Padding color
Quality = Compression quality
Sharpen = Sharpen the image


Photo List File Example:

<photolist count="*" sort="random" desc=1 repeat="0">
    <imagefile>c:\foo.jpg</imagefile>
    <imagefile>http://www.msn.com/foo.jpg</imagefile>
    <photolist count= sort= desc= repeat=>
        ....  
    </photolist>
    <photolistfile count="5" sort="datem" desc="1" repeat="1">
        c:\photolist.xml
    </photolistfile>
    <directory count="*" sort="datec" desc="0" repeat="0" dirmask="*" filemask="*.jpg" recurse="1" sort="name" desc="0" root="even">
        c:\foo\
    </directory>
    <group count="3" sort="random" desc= repeat= dirmask= filemask= recurse= sort= desc= root= depth=>
        http://groups.msn.com/mygroup
    </group>
</photolist>

A photo list is a command to describe a (potentially recursive) set of photos to select.  The top level commands are:

imagefile = Select this image file from a URL or file name
photolist = Execute the embedded list and return the photos this list selects
photolistfile = Execute the specified photo list file.  Any attributes in the top level command in photolist file are ignored
directory = Select photos (potentially recursively) from a directory applying a normalization and selection within interior nodes
group = Select photos (potentially recursively) from an MSN Group applying a normalization and selection within interior nodes

Commands/Attributes common to photolist, photolistfile, directory, and group:

count = Number of photos to select from the list.  If the parameter is "*" it selects all photos.  The default is "*".
sort = Sort order of the photos.  Values can be "random" (photos randomly selected), "datem" (date modified), "datec" (date created), "size" (size of the photo), "name" (name of the file), or "listed" (specified order from list).  The default is "listed".
desc = Sort descending if "1".  Otherwise ascending sort.  Default is "1"
repeat = Repeat photos to fill count if the number of photos in the list is less than the count if "1".  Default is "1".

Commands unique to directory and group:

dirmask = Mask to filter directories selected.  Default is "*".
filemask = Mask to filter files selected.  Default is "*.jpg".
root = Controls how a "root" or "interior" node's files are treated.  If "none", files in a directory with sub-directories are ignored.  If "even", files in a directory are treated on parity with files in sub-directory. In other words, consider the files in a directory all moved to a sub-directory and treated the same as sibling directories.  If "heavy", files in a directory are considered after selecting photos from all the sub-directires and treated as parity with this selection.
depth = Controls the depth at which normalization occurs.  Normalization and selection according to the parameters will only occur when the sub-directory is deeper than the specified depth.  

Consider the following example:

a\one.jpg
a\two.jpg
a\b\c\three.jpg
a\b\c\four.jpg
a\b\d\five.jpg
a\b\d\six.jpg

With a depth of 3, a count of 1, a sort of random, and a root of even the directory command would select the files one.jpg, two.jpg, a random file from three/four.jpg and a random file from  five/six.jpg.  So an example returned might be one.jpg, two.jpg, three.jpg, six.jpg.


