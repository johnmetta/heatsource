# Heat Source
Heat Source is a computer model used by Oregon DEQ to simulate stream thermodynamics and hydraulics.

## Development History

### Original Version
Heat Source was developed in 1996 by Matt Boyd as a Masters Thesis at Oregon State University in 
the Departments of Bioresource Engineering and Civil Engineering and programmed in Visual Basic as 
an Excel macro. 

The original programming can be found in the Excel file found in the attic. There is
also a PDF of the original manual which discusses the methodology and mathematics and a copy of
"Shade-a-lator," a shade calculation algorithm that was directly incorporated into the current 
version.

### Current Version
In 2006, I (John Metta) and Dan Turner redeveloped it from the ground up as a Python (V2.5)library 
while we were both working for the DEQ Watershed Management division. This source is what you will 
find in the src directory.

### Science and Mathematics
Peer review documents regarding this software are available at: 
http://www.deq.state.or.us/wq/TMDLs/heatsource.htm

## Fine Print
DEQ currently maintains the official Heat Source methodology and source. Details are available at 
http://www.deq.state.or.us/wq/TMDLs/tools.htm. This version is a branch that I will be building from
as I have time. Much of what I will work on will become separate, stand-alone libraries (e.g. solar
position will be separated from hydraulic, etc.). 

There is the possibility that I will merge some of these changes as appropriate with the official DEQ
version through an ongoing collaboration with the DEQ Watershed Management Division. This will only be
for small changes that will keep the DEQ version much as it currently is (i.e. backward compatible).

## License

The DEQ version of this software was licensed as Public Domain. This license holds for the initial 
import of this source tree. Changes from the initial source tree are licensed as follows:

Copyright (c) 2009 John Metta

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.