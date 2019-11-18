from psd_tools import PSDImage
import os

srcFolder = "H:\\赋能起航志愿者证书\\November\\PSD\\"
dstFolder = "H:\\赋能起航志愿者证书\\November\\JPG\\"

for file in os.listdir(srcFolder):
	srcFile = os.path.join(srcFolder, file)
	if os.path.isfile(srcFile) == True:
		psdFile = PSDImage.open(srcFile)
		dstFile = os.path.join(dstFolder, os.path.splitext(file)[0] + ".jpg")
		psdFile.compose().save(dstFile)

print("Finished!")


