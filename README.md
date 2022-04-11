# PSWriteWord - PowerShell Module

<p align="center">
  <a href="https://dev.azure.com/evotecpl/PSWriteWord/_build/results?buildId=latest"><img src="https://dev.azure.com/evotecpl/PSWriteWord/_apis/build/status/EvotecIT.PSWriteWord"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteWord"><img src="https://img.shields.io/powershellgallery/v/PSWriteWord.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteWord"><img src="https://img.shields.io/powershellgallery/vpre/PSWriteWord.svg?label=powershell%20gallery%20preview&colorB=yellow"></a>
  <a href="https://github.com/EvotecIT/PSWriteWord"><img src="https://img.shields.io/github/license/EvotecIT/PSWriteWord.svg"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSWriteHTML"><img src="https://img.shields.io/powershellgallery/p/PSWriteWord.svg"></a>
  <a href="https://github.com/EvotecIT/PSWriteWord"><img src="https://img.shields.io/github/languages/top/evotecit/PSWriteWord.svg"></a>
  <a href="https://github.com/EvotecIT/PSWriteWord"><img src="https://img.shields.io/github/languages/code-size/evotecit/PSWriteWord.svg"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteWord"><img src="https://img.shields.io/powershellgallery/dt/PSWriteWord.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

**New version of this module will be developed under [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) project.**
***This module will no longer be developed!*** It has been replaced by [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) project which will be combination of Word, Excel and in future PowerPoint features. No new features, fixes will be added, but it will continue to work. Due to license changes to DLL it's not possible to support this project. **PSWriteOffice** is complete rewrite and cross-platform.

***

_PSWriteWord_ is PowerShell module to create Microsoft Word documents without Microsoft Word installed. You can read some more information on my [website](https://evotec.xyz/hub/scripts/pswriteword-powershell-module/).

## Example usage of Add-WordTable / Add-WordText in action

![image](https://evotec.xyz/wp-content/uploads/2018/07/PSWriteWord-MoreAction.gif)

## Requirements

- Works only on Windows (as NET CORE is not supported by DLL) - until Xceed ads that to their version. Which will take a while for the free version to have (if ever).
- Works only on PowerShell 5.1

## Systems it should run on (marked those confirmed to run)

### Windows Client Systems

- ☑ Windows 7 with Service Pack 1 - requires WMF 5.1 <https://www.microsoft.com/en-us/download/details.aspx?id=54616>
- not Windows 8.1
- ☑ Windows 10 1809
- ☑ Windows 10 1803
- ☑ Windows 10 1809
- ☑ Windows 10 1903
- ☑ Windows 10 1909

### Windows Server Systems

- ☑ Windows 2008 R2 with Service Pack 1 - requires WMF 5.1 <https://www.microsoft.com/en-us/download/details.aspx?id=54616>
- ☑ Windows 2012 R2
- ☑ Windows 2016

## Credits

This module is based on **DocX** from **Xceed Software** (<https://github.com/xceedsoftware/DocX>). DocX is the free, open source version of **Xceed Words for .NET**. Originally written by Cathal Coffey, and maintained by Przemyslaw Klys (me), it is now maintained by **Xceed**. This also means any bugs / issues with C# version of it will be available in **PSWriteWord**. On the other hand since Xceed releases new versions of DocX (just a few versions behind the pay version) it means eventually **PSWriteWord** will get even more features. This also means that if something is not yet available in commands that are listed below or are incomplete you can take the .net approach and simply use it that way.

### License

**MIT License** applies only to PowerShell code/code within this repo. DLL is property of Xceed and as such is licensed under **Microsoft Public License (Ms-PL)**.
