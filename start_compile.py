from support.office import powerpoint

addin_name = "PowerPoint_tools"
powerpoint.unregister_addin(addin_name)
powerpoint.compile_addin(addin_name + ".ppam", "manifest.py", keep_pptm=False)
powerpoint.register_addin(addin_name + ".ppam")
