
IfsToolboxps.dll: dlldata.obj IfsToolbox_p.obj IfsToolbox_i.obj
	link /dll /out:IfsToolboxps.dll /def:IfsToolboxps.def /entry:DllMain dlldata.obj IfsToolbox_p.obj IfsToolbox_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del IfsToolboxps.dll
	@del IfsToolboxps.lib
	@del IfsToolboxps.exp
	@del dlldata.obj
	@del IfsToolbox_p.obj
	@del IfsToolbox_i.obj
