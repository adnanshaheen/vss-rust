use std::ptr;
use winapi::shared::guiddef::*;
use winapi::shared::winerror::S_OK;
use winapi::shared::winerror::*;
use winapi::um::combaseapi::StringFromGUID2;
use winapi::um::combaseapi::*;
use winapi::um::objbase::*;
use winapi::um::vsbackup::*;
use winapi::um::vss::*;
use winapi::um::winbase::INFINITE;

fn hresult_to_hex(hr: i32) -> String {
    format!("0x{:08X}", hr) // Format the HRESULT as an 8-digit hexadecimal
}

/// Converts a GUID to a String
fn guid_to_string(guid: &GUID) -> String {
    let mut buffer: [u16; 39] = [0; 39]; // GUID string is 38 chars + null terminator

    unsafe {
        let len = StringFromGUID2(guid, buffer.as_mut_ptr(), buffer.len() as i32);
        if len > 0 {
            // Convert wide-char buffer to String and trim null terminators
            String::from_utf16_lossy(&buffer)
                .trim_end_matches('\0')
                .to_string()
        } else {
            String::from("Conversion failed")
        }
    }
}

/// Struct to hold details of VSS writers
struct WriterDetails {
    pub writer_id: String,
    pub writer_name: String,
}

fn wait_for_async(p_async: *mut IVssAsync) -> HRESULT {
    unsafe {
        if p_async.is_null() {
            eprintln!("VSS async is null.");
            return E_POINTER;
        }

        // Call Wait() to wait for the operation to complete
        let mut hr = (*p_async).Wait(INFINITE);
        if FAILED(hr) {
            eprintln!("Wait failed with error: {}", hresult_to_hex(hr));
            return hr;
        }

        // Query the status of the async operation
        let mut hr_status: HRESULT = 0;
        hr = (*p_async).QueryStatus(&mut hr_status, ptr::null_mut());

        if FAILED(hr_status) {
            eprintln!("QueryStatus failed with error: {}", hresult_to_hex(hr));
            return hr_status;
        }

        // Return the final operation status
        eprintln!("Result: {}", hresult_to_hex(hr_status));
        hr_status
    }
}

fn list_vss_writers() -> Vec<WriterDetails> {
    unsafe {
        // Initialize COM
        let mut hr = CoInitializeEx(ptr::null_mut(), COINIT_APARTMENTTHREADED);
        if hr != S_OK {
            eprintln!("CoInit failed with error: {}", hresult_to_hex(hr));
            return Vec::new();
        }

        // Create VSS Backup Components
        let mut p_vss: *mut IVssBackupComponents = ptr::null_mut();
        hr = CreateVssBackupComponents(&mut p_vss);
        if hr != S_OK || p_vss.is_null() {
            CoUninitialize();
            eprintln!(
                "CreateVssBackupComponents failed with error: {}",
                hresult_to_hex(hr)
            );
            return Vec::new();
        }

        // Initialize the backup components
        hr = (*p_vss).InitializeForBackup(ptr::null_mut());
        if hr != S_OK {
            (*p_vss).Release();
            CoUninitialize();
            eprintln!(
                "Initialize for backup failed with error: {}",
                hresult_to_hex(hr)
            );
            return Vec::new();
        }

        hr = (*p_vss).SetBackupState(false, true, VSS_BT_FULL, false);
        if FAILED(hr) {
            (*p_vss).Release();
            CoUninitialize();
            eprintln!(
                "Failed to set backup state with error: {}",
                hresult_to_hex(hr)
            );
            return Vec::new();
        }

        // Gather writer metadata
        let mut m_vss_sync: *mut IVssAsync = ptr::null_mut();
        hr = (*p_vss).GatherWriterMetadata(&mut m_vss_sync);
        if hr != S_OK {
            (*p_vss).Release();
            CoUninitialize();
            eprintln!(
                "GatherWriterMetadata failed with error: {}",
                hresult_to_hex(hr)
            );
            return Vec::new();
        }

        // Wait for operation to complete
        hr = wait_for_async(m_vss_sync);
        if FAILED(hr) {
            (*p_vss).Release();
            CoUninitialize();
            eprintln!("wait_for_async failed with error: {}", hresult_to_hex(hr));
            return Vec::new();
        }

        // Get writer status count
        let mut writer_count = 0;
        hr = (*p_vss).GetWriterStatusCount(&mut writer_count);
        if FAILED(hr) {
            (*p_vss).Release();
            CoUninitialize();
            eprintln!(
                "GetWriterStatusCount failed with error: {}",
                hresult_to_hex(hr)
            );
            return Vec::new();
        }

        let mut writers = Vec::new();
        eprintln!("Number of writers: {}", writer_count);

        // Loop through each writer and extract details
        for i in 0..writer_count {
            let mut instance_id = GUID {
                Data1: 0,
                Data2: 0,
                Data3: 0,
                Data4: [0; 8],
            };
            let mut writer_id = GUID {
                Data1: 0,
                Data2: 0,
                Data3: 0,
                Data4: [0; 8],
            };
            let mut writer_name: *mut u16 = ptr::null_mut();
            let mut state = 0;
            let mut failure_reason = 0;

            if (*p_vss).GetWriterStatus(
                i,
                &mut instance_id as *mut _,
                &mut writer_id as *mut _,
                &mut writer_name,
                &mut state,
                &mut failure_reason,
            ) == S_OK
            {
                let writer_id_str = guid_to_string(&writer_id);
                let writer_name_str = format!("{:?}", writer_name);

                writers.push(WriterDetails {
                    writer_id: writer_id_str,
                    writer_name: writer_name_str,
                });
            }
        }

        // Free writer status resources
        (*p_vss).FreeWriterStatus();
        (*p_vss).Release();
        CoUninitialize();

        writers
    }
}

fn main() {
    let writers = list_vss_writers();
    println!("List of VSS Writers:");
    for writer in writers {
        println!("Id: {}", writer.writer_id);
        println!("Name: {}", writer.writer_name);
    }
}
