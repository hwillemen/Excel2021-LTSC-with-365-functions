# Introduction 
This repo contains VBA code for Excel2021 to integrate some very useful Excel365 functions.<br/>
The worksheet contains the CodeModule in the VBA editor and also has sheets to test the different corner cases of these functions.<br/>
Definitely check out these tests if you require a more exotic parameter/feature/behavior.<br/>
Excel365-func.vbs contains a copy of the code for easy access and so that Git systems detect the vbs language.

# support: 
- TEXTSPLIT -> only ignore_empty is not yet supported
- TEXTBEFORE -> Fully supported but behaves different (more logical?) if delimiter is not found with match_end
- TEXTAFTER -> Fully supported but behaves different (more logical?) if delimiter is not found with match_end
- TOCOL -> Fully supported
- TOROW -> Fully supported
- VSTACK -> Fully supported
- HSTACK -> Fully supported
- WRAPROWS -> No plan (yet) to support
- WRAPCOLS -> No plan (yet) to support
- TAKE -> No plan (yet) to support
- DROP -> No plan (yet) to support
- CHOOSECOLS -> No plan (yet) to support
- CHOOSEROWS -> No plan (yet) to support
- EXPAND -> No plan (yet) to support

Other useful functions:
- ArraySize -> checks if input is an array and if so returns the size. A Range is not 100% compatible so throws a specific error.
- ToArray -> converts the input to an array, returns the 1st size like ArraySize() but also ArrayInfo for 1st and 2nd dimension.
- ResizeArray -> Works like a Redim Preserve but on all dimensions, not just 1 like Redim Preserve<br/>See the TOCOL,TOROW,... sheet for examples.
- SortUnique -> WIP: for conveniant filtering on blanks, on unique and sorting
