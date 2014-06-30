class ZipFile
{
	pack(src, dest, del:=false) {
		fso := ComObjCreate("Scripting.FileSystemObject")
		, psh := ComObjCreate("Shell.Application")
		, src := fso.GetAbsolutePathName(src)
		, dest := fso.GetAbsolutePathName(dest)

		if !FileExist(dest) {
			header1 := "PK" . Chr(5) . Chr(6)
			, VarSetCapacity(header2, 18, 0)
			, archive := FileOpen(dest, "w")
			, archive.Write(header1)
			, archive.RawWrite(header2, 18)
			, archive.Close()
		}

		if !(pzip:=psh.NameSpace(dest))
			throw Exception("Failed to create folder object.", -1)
		items := psh.NameSpace(fso.GetParentFolderName(src)).Items()
		items.Filter(0x00020|0x00040|0x00080, fso.GetFileName(src))
		for item in items {
			pzip[del ? "MoveHere" : "CopyHere"](item, 4|16)
			i := A_Index, done := -1
			while (done != i)
				done := pzip.Items().Count
		}
		return done ;// return total number of zipped items
	}

	unpack(src, dest:="", del:=false) {
		fso := ComObjCreate("Scripting.FileSystemObject")
		, psh := ComObjCreate("Shell.Application")
		, src := fso.GetAbsolutePathName(src)
		, dest := fso.GetAbsolutePathName(dest)

		;// create temporary destination folder if 'dest' exists
		if (temp := fso.FolderExists(dest))
			_prev := psh.NameSpace(dest)
			, dest .= "\" . fso.GetTempName()
		fso.CreateFolder(dest)
		, _src := psh.NameSpace(src)
		, _dest := psh.NameSpace(dest)
		, _dest.CopyHere(_src.Items(), 4|16)

		zipped := _src.Items().Count
		while (_dest.Items().Count != zipped)
			Sleep 15

		if temp {
			for item in _dest.Items() {
				_prev.MoveHere(item, 4|16)
				while (_dest.Items().Count != 0)
					Sleep 15
			}
			fso.DeleteFolder(dest, true)
		}
		if del
			FileDelete %src%
	}
}