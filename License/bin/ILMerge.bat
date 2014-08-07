ECHO parameter=%1
CD %1
COPY License.exe temp.exe
ILMerge.exe /targetplatform:v4 /out:License.exe temp.exe CMS.dll
DEL temp.exe
DEL CMS.*
DEL License.pdb