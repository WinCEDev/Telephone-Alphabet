Attribute VB_Name = "ImageListEx"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bitmap load functions.

Public Declare Function ImageListEx_SHLoadDIBitmap _
               Lib "Coredll" _
               Alias "SHLoadDIBitmap" (ByVal szFileName As String) As Long

Public Declare Function ImageListEx_DeleteObject _
               Lib "Coredll" _
               Alias "DeleteObject" (ByVal hObject As Long) As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ImageList functions.

Public Declare Function ImageListEx_ImageList_Create _
               Lib "Coredll" _
               Alias "ImageList_Create" (ByVal cx As Long, _
                                         ByVal cy As Long, _
                                         ByVal Flags As Long, _
                                         ByVal cInitial As Long, _
                                         ByVal cGrow As Long) As Long
               
Public Declare Function ImageListEx_ImageList_Add _
               Lib "Coredll" _
               Alias "ImageList_Add" (ByVal himl As Long, _
                                      ByVal hbmImage As Long, _
                                      ByVal hbmMask As Long) As Long

Public Declare Function ImageListEx_ImageList_AddMasked _
               Lib "Coredll" _
               Alias "ImageList_AddMasked" (ByVal himl As Long, _
                                            ByVal hbmImage As Long, _
                                            ByVal crMask As Long) As Long

Public Declare Function ImageListEx_ImageList_Replace _
               Lib "Coredll" _
               Alias "ImageList_Replace" (ByVal himl As Long, _
                                          ByVal i As Long, _
                                          ByVal hbmImage As Long, _
                                          ByVal hbmMask As Long) As Long

Public Declare Function ImageListEx_ImageList_GetImageCount _
               Lib "Coredll" _
               Alias "ImageList_GetImageCount" (ByVal himl As Long) As Long

Public Declare Function ImageListEx_ImageList_Remove _
               Lib "Coredll" _
               Alias "ImageList_Remove" (ByVal himl As Long, _
                                         ByVal i As Long) As Boolean

Public Declare Function ImageListEx_ImageList_RemoveAll _
               Lib "Coredll" _
               Alias "ImageList_RemoveAll" (ByVal himl As Long) As Boolean

Public Declare Function ImageListEx_ImageList_Destroy _
               Lib "Coredll" _
               Alias "ImageList_Destroy" (ByVal himl As Long) As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ImageList flag values.
'https://learn.microsoft.com/en-us/previous-versions/ms960944(v=msdn.10)
Public Const ILC_COLOR    As Long = &H0

Public Const ILC_COLORDDB As Long = &HFE

Public Const ILC_MASK     As Long = &H1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mask values.

Private Const CLR_DEFAULT As Long = &HFF000000

Public Function ImageListEx_Replace(ByVal ImageListHandle As Long, _
                                    ByVal Index As Long, _
                                    ByVal ImagePath As String) As Long
Attribute ImageListEx_Replace.VB_Description = "Replaces an image in an image list with a new image. Returns nonzero if successful, or zero otherwise."

    Dim lngBitmap As Long

    lngBitmap = ImageListEx_SHLoadDIBitmap(ImagePath)

    ImageListEx_Add = ImageListEx_ImageList_Replace(ImageListHandle, Index, ImagePath, 0)

    ImageListEx_DeleteObject lngBitmap

End Function

Public Function ImageListEx_Remove(ByVal ImageListHandle As Long, _
                                   ByVal Index As Long) As Boolean
Attribute ImageListEx_Remove.VB_Description = "Removes an image from an image list. Returns True if successful, or False otherwise."
    ImageListEx_Remove = ImageListEx_ImageList_Remove(ImageListHandle, Index)
End Function

Public Function ImageListEx_RemoveAll(ByVal ImageListHandle As Long) As Boolean
Attribute ImageListEx_RemoveAll.VB_Description = "Removes all of the images from an image list. Returns nonzero if successful, or zero otherwise."
    ImageListEx_RemoveAll = ImageListEx_ImageList_RemoveAll(ImageListHandle)
End Function

Public Function ImageListEx_Add(ByVal ImageListHandle As Long, _
                                ByVal ImagePath As String) As Long
Attribute ImageListEx_Add.VB_Description = "Adds an image or images to an image list. Returns the index of the first new image if successful, or -1 otherwise."

    Dim lngBitmap As Long

    lngBitmap = ImageListEx_SHLoadDIBitmap(ImagePath)

    ImageListEx_Add = ImageListEx_ImageList_Add(ImageListHandle, lngBitmap, 0)

    ImageListEx_DeleteObject lngBitmap

End Function

Public Function ImageListEx_AddMasked(ByVal ImageListHandle As Long, _
                                      ByVal ImagePath As String, _
                                      ByVal MaskColor As ColorConstants) As Long
Attribute ImageListEx_AddMasked.VB_Description = "Adds an image or images to an image list. Specify a color to use as the image mask, or if this parameter is CLR_DEFAULT, then the color of the pixel at (0,0) is used as the mask. Returns the index of the first new image if successful, or -1 otherwise."

    Dim lngBitmap As Long

    lngBitmap = ImageListEx_SHLoadDIBitmap(ImagePath)

    ImageListEx_AddMasked = ImageListEx_ImageList_AddMasked(ImageListHandle, lngBitmap, MaskColor)

    ImageListEx_DeleteObject lngBitmap

End Function

Public Function ImageListEx_Create(ByVal ImageWidth As Long, _
                                   ByVal ImageHeight As Long, _
                                   ByVal Flags As Long) As Long
Attribute ImageListEx_Create.VB_Description = "Creates a new image list. Returns the handle to the image list if successful, or 0 otherwise."

    Const INITIAL_IMAGES   As Long = 0 'How many images the ImageList shoud initially contain.

    Const GROW_BY          As Long = 1 'By how many images to grow the ImageList when a new image is added.

    Dim lngImageListHandle As Long

    lngImageListHandle = ImageListEx_ImageList_Create(ImageWidth, ImageHeight, Flags, INITIAL_IMAGES, GROW_BY)

    ImageListEx_Create = lngImageListHandle
End Function

Public Function ImageListEx_Destroy(ByVal ImageListHandle As Long) As Long
Attribute ImageListEx_Destroy.VB_Description = "Destroys an image list. Returns nonzero if successful, or zero otherwise."
    ImageListEx_Destroy = ImageListEx_ImageList_Destroy(ImageListHandle)
End Function

Public Function ImageListEx_GetImageCount(ByVal ImageListHandle As Long) As Long
Attribute ImageListEx_GetImageCount.VB_Description = "Retrieves the number of images in an image list. Returns the number of images if successful."
    ImageListEx_GetImageCount = ImageListEx_ImageList_GetImageCount(ImageListHandle)
End Function

