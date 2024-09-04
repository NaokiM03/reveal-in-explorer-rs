use std::{io, path::Path};

use windows::{
    core::HSTRING,
    Win32::{
        System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED},
        UI::Shell::{Common::ITEMIDLIST, ILCreateFromPathW, ILFree, SHOpenFolderAndSelectItems},
    },
};

pub fn reveal_in_explorer(path: impl AsRef<Path>) -> io::Result<()> {
    unsafe {
        let pvreserved = None;
        let dwcoinit = COINIT_MULTITHREADED;
        CoInitializeEx(pvreserved, dwcoinit).ok()?;

        let pidlfolder: *const ITEMIDLIST = {
            let path = path.as_ref();
            let pszpath = HSTRING::from(path);
            ILCreateFromPathW(&pszpath)
        };

        if pidlfolder.is_null() {
            return Err(io::Error::last_os_error());
        }

        let result = {
            let apidl = None;
            let dwflags = 0;
            SHOpenFolderAndSelectItems(pidlfolder, apidl, dwflags)
        };

        let pidl = Some(pidlfolder);
        ILFree(pidl);

        CoUninitialize();

        result?;
    }

    Ok(())
}
