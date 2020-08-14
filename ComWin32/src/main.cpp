#include <windows.h>
#include <shobjidl.h> 

// Component Object Model (COM)

/*
-COM uses reference counting to free resources, every COM object maintains an internal counter
-the internal counter comes from inheriting the IUnknown interface (all COM interfaces must inherit from the IUnknown interface)
-IUnknown interface defines three methods; 
    - QueryInterface -> enables a program to query the capabilities of the object at run time
    - AddRef & Release -> used to control the lifetime of an object 
*/

template <class T> 
void SafeRelease(T** ppT)
{
    /*
    - allows to safely release an interface pointer whether or not it is valid 
    - This function takes a COM interface pointer as a parameter and does the following:
        - Checks whether the pointer is NULL
        - Calls Release if the pointer is not NULL
        - Sets the pointer to NULL
    */
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = NULL;
    }
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR pCmdLine, int nCmdShow)
{
    // Initialize the COM library - Apartment threaded  
    // HRESULT contains a success or error code  
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED |
        COINIT_DISABLE_OLE1DDE);
    
    if (SUCCEEDED(hr)) // alternatively can use the 'FAILED' macro
    {
        // have to call 'Release' on 'pFileOpen' pointer to decrement reference counter, if successful
        IFileOpenDialog* pFileOpen;

        // Create the FileOpenDialog object - returns a success or error code 
        hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
            IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));

        if (SUCCEEDED(hr))
        {
            // Show the Open dialog box - blocks until the user dismisses the dialog box 
            hr = pFileOpen->Show(NULL);

            // Get the file name from the dialog box.
            if (SUCCEEDED(hr))
            {
                
                // Shell item object, represents the file the user selected  
                // have to call 'Release' on 'pItem' pointer to decrement reference counter, if successful
                IShellItem* pItem; 
                hr = pFileOpen->GetResult(&pItem);
                if (SUCCEEDED(hr))
                {
                    PWSTR pszFilePath;
                    hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath); // get the file path in the form of a string 

                    // Display the file name to the user
                    if (SUCCEEDED(hr))
                    {
                        MessageBox(NULL, pszFilePath, L"File Path Selected", MB_OK);
                        
                        // frees a block of memory that was allocated with 'CoTaskMemAlloc'
                        // in this case, 'GetDisplayName' allocates memory for a string ('CoTaskMemAlloc' is called internally)
                        CoTaskMemFree(pszFilePath);
                    }
                    pItem->Release();
                }
            }
        }

        SafeRelease(&pFileOpen);
        
        // must call before the thread exits - takes no paramters and has no return value 
        CoUninitialize();
    }
    return 0;
}