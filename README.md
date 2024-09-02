# WeChatOCR
这是一个采用Python调用微信OCR功能，进行批处理图片OCR的代码。

首先非常感谢swigger，52PJ的FeiyuYip，nulptr以及其他对此做出贡献的朋友。

基于他们的工作，改动如下：
1. 将WeChatOCR.exe做了本地化，不再依赖微信的安装路径。
2. 将图片处理的格式多样化，增加了jpg，jpeg，bmp，tif格式的处理，只需要将文件放入scr文件夹中的即可。
3. 将OCR的处理结果将以docx格式保存到docx文件夹中。

关于源文件的问题：
我感觉wenchatocr对png格式的处理能力比较好，所以建议将图片格式转换为png以后再做OCR处理。
