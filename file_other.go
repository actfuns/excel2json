//go:build !windows

package main

import "os"

// openFileShared opens a file for reading with shared access.
// On Unix, os.Open already permits concurrent access.
func openFileShared(path string) (*os.File, error) {
	return os.Open(path)
}
