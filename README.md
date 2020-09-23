<div align="center">

## Alpha gradient Pictures

<img src="PIC20091030739218505.JPG">
</div>

### Description

This code draws two pictures that merged - with gradually alpha pixels (like in photoshop)this code is realy wonderful but it's not so hard although its fast, it uses AlphaBlend API function, that requires the destination hDC and source hDC and both's recangle (left,top,width,height) and the blend value(beware: the blend value is between 0 and 255, but in AlphaBlend API function multiply the blend value by 256^2 or in hex add 4 zeroez in the right side like: &amp;H9F --&gt; &amp;H9F0000)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2009-10-22 22:28:50
**By**             |[Saed abumokh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/saed-abumokh.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Alpha\_grad21665010302009\.zip](https://github.com/Planet-Source-Code/saed-abumokh-alpha-gradient-pictures__1-72606/archive/master.zip)

### API Declarations

```
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal
```





