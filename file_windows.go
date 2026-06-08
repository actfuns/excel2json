//go:build windows

package main

import (
	"os"
	"syscall"
)

// openFileShared opens a file for reading with FILE_SHARE_READ | FILE_SHARE_WRITE,
// matching C#'s FileShare.ReadWrite so that Excel can keep the file open while we read it.
func openFileShared(path string) (*os.File, error) {
	pathp, err := syscall.UTF16PtrFromString(path)
	if err != nil {
		return nil, err
	}
	handle, err := syscall.CreateFile(
		pathp,
		syscall.GENERIC_READ,
		uint32(syscall.FILE_SHARE_READ|syscall.FILE_SHARE_WRITE),
		nil,
		syscall.OPEN_EXISTING,
		syscall.FILE_ATTRIBUTE_NORMAL,
		0,
	)
	if err != nil {
		return nil, err
	}
	return os.NewFile(uintptr(handle), path), nil
}
