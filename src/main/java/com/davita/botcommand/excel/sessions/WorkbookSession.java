package com.davita.botcommand.excel.sessions;

import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.EncryptedDocumentException;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.channels.FileChannel;
import java.nio.channels.FileLock;
import java.nio.channels.OverlappingFileLockException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;


/**
 * Windows-optimized session for managing Excel workbooks with Apache POI.
 * Uses File-based WorkbookFactory for reduced memory consumption and
 * RandomAccessFile for Windows file locking.
 *
 * Designed exclusively for Windows environments with Automation Anywhere A360.
 */
public class WorkbookSession implements CloseableSessionObject {

    private static final long LOCK_TIMEOUT_MS = 10_000L; // 10 seconds
    private static final long LOCK_RETRY_INTERVAL_MS = 100L;
    private static final String TEMP_SUFFIX = ".tmp";

    private volatile boolean closed = false;
    private Workbook workbook;
    private File file;
    private boolean readOnly;

    // Windows file locking via RandomAccessFile
    private RandomAccessFile randomAccessFile;
    private FileChannel fileChannel;
    private FileLock fileLock;

    // ========== CONSTRUCTORS ==========

    /**
     * Private constructor for internal use by factory methods.
     */
    private WorkbookSession() {
    }

    // ========== PUBLIC STATIC FACTORY METHODS ==========

    /**
     * Creates a new Excel workbook file with the specified format.
     * Automatically determines format based on file extension (.xls or .xlsx).
     *
     * For Windows environments, this method:
     * - Creates parent directories if needed
     * - Initializes the workbook with one sheet
     * - Acquires an exclusive file lock
     * - Does NOT write to disk until save() is called
     *
     * @param filePath The absolute path where the workbook will be saved (must end with .xls or .xlsx)
     * @param initialSheetName Optional name for the first sheet (defaults to "Sheet1" if null/empty)
     * @return A new WorkbookSession ready for data manipulation
     * @throws IOException if file path is invalid, format is unsupported, or file system error occurs
     */
    public static WorkbookSession createWorkbook(String filePath, String initialSheetName) throws IOException {
        validateFilePath(filePath);

        File targetFile = new File(filePath);
        Path targetPath = targetFile.toPath().toAbsolutePath();

        // Create parent directories
        Path parent = targetPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        // Determine workbook type from extension
        String fileName = targetFile.getName().toLowerCase();
        Workbook workbook;

        if (fileName.endsWith(".xls")) {
            workbook = new HSSFWorkbook();
        } else if (fileName.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (fileName.endsWith(".xlsm")) {
            workbook = new XSSFWorkbook();
        } else {
            throw new IOException("E-EXT-INVALID: Unsupported file extension. Must be .xls, .xlsx, or .xlsm. Got: "
                    + fileName);
        }

        // Create initial sheet
        String sheetName = (initialSheetName == null || initialSheetName.trim().isEmpty())
                ? "Sheet1"
                : initialSheetName.trim();
        workbook.createSheet(sheetName);

        // Write empty file so it exists on disk
        // This is needed for file locking to work
        if (!targetFile.exists()) {
            try (FileOutputStream fos = new FileOutputStream(targetFile)) {
                // Write empty file or minimal content
                fos.flush();
            }
        }

        // Initialize session WITHOUT reopening the file
        // Keep the workbook in memory only
        WorkbookSession session = new WorkbookSession();
        session.workbook = workbook;
        session.file = targetFile;
        session.readOnly = false;

        // Acquire exclusive lock on the empty file
        session.acquireLock();

        return session;
    }

    /**
     * Opens an existing Excel workbook from the specified file path.
     * Uses WorkbookFactory.create(File) for optimal memory efficiency.
     *
     * This method provides the lowest memory footprint for reading Excel files
     * by using File objects instead of InputStreams, which is critical for
     * large workbooks in Windows environments.
     *
     * @param filePath The absolute path to the existing workbook file
     * @param readOnly If true, opens with shared lock (read-only); if false, exclusive lock (read-write)
     * @return A WorkbookSession containing the loaded workbook
     * @throws IOException if file doesn't exist, is not a valid Excel file, or locking fails
     * @throws EncryptedDocumentException if the workbook is password-protected
     */
    public static WorkbookSession openWorkbook(String filePath, boolean readOnly) throws IOException {
        validateFilePath(filePath);

        File targetFile = new File(filePath);

        if (!targetFile.exists()) {
            throw new IOException("E-FILE-NOTFOUND: File does not exist: " + targetFile.getAbsolutePath());
        }

        if (!targetFile.isFile()) {
            throw new IOException("E-NOT-FILE: Path is not a regular file: " + targetFile.getAbsolutePath());
        }

        if (!targetFile.canRead()) {
            throw new IOException("E-NO-READ-ACCESS: Cannot read file: " + targetFile.getAbsolutePath());
        }

        // Initialize session
        WorkbookSession session = new WorkbookSession();
        session.file = targetFile;
        session.readOnly = readOnly;

        try {
            // WorkbookFactory automatically handles both .xls and .xlsx
            session.workbook = WorkbookFactory.create(targetFile);

            // Acquire lock AFTER workbook is loaded
            session.acquireLock();

        } catch (EncryptedDocumentException e) {
            throw new EncryptedDocumentException("E-ENCRYPTED: Workbook is password-protected: "
                    + targetFile.getAbsolutePath() + ". Use openWorkbook(filePath, password, readOnly) for encrypted files.", e);
        } catch (IOException e) {
            // Clean up workbook if lock acquisition failed
            if (session.workbook != null) {
                try {
                    session.workbook.close();
                } catch (Exception ignored) {}
            }
            throw new IOException("E-OPEN-FAIL: Failed to open workbook from " + filePath
                    + ". Ensure the file is a valid Excel file (.xls/.xlsx): " + e.getMessage(), e);
        } catch (Exception e) {
            // Clean up workbook if unexpected error
            if (session.workbook != null) {
                try {
                    session.workbook.close();
                } catch (Exception ignored) {}
            }
            throw new IOException("E-UNEXPECTED: Unexpected error opening workbook: " + e.getMessage(), e);
        }

        return session;
    }

    /**
     * Opens an existing password-protected Excel workbook.
     * Uses WorkbookFactory.create(File, password) for memory-efficient loading.
     *
     * @param filePath The absolute path to the existing workbook file
     * @param password The password to decrypt the workbook
     * @param readOnly If true, opens with shared lock (read-only); if false, exclusive lock (read-write)
     * @return A WorkbookSession containing the loaded workbook
     * @throws IOException if file doesn't exist, password is incorrect, or locking fails
     */
    public static WorkbookSession openWorkbook(String filePath, String password, boolean readOnly) throws IOException {
        validateFilePath(filePath);

        File targetFile = new File(filePath);

        if (!targetFile.exists()) {
            throw new IOException("E-FILE-NOTFOUND: File does not exist: " + targetFile.getAbsolutePath());
        }

        if (!targetFile.isFile()) {
            throw new IOException("E-NOT-FILE: Path is not a regular file: " + targetFile.getAbsolutePath());
        }

        // Initialize session
        WorkbookSession session = new WorkbookSession();
        session.file = targetFile;
        session.readOnly = readOnly;

        try {
            // Load workbook FIRST
            if (password != null && !password.isEmpty()) {
                session.workbook = WorkbookFactory.create(targetFile, password);
            } else {
                session.workbook = WorkbookFactory.create(targetFile);
            }

            // THEN acquire lock
            session.acquireLock();

        } catch (EncryptedDocumentException e) {
            throw new IOException("E-WRONG-PASSWORD: Incorrect password or workbook encryption error: "
                    + e.getMessage(), e);
        } catch (IOException e) {
            // Clean up workbook if lock acquisition failed
            if (session.workbook != null) {
                try {
                    session.workbook.close();
                } catch (Exception ignored) {}
            }
            throw new IOException("E-OPEN-FAIL: Failed to open encrypted workbook from " + filePath
                    + ": " + e.getMessage(), e);
        } catch (Exception e) {
            // Clean up workbook if unexpected error
            if (session.workbook != null) {
                try {
                    session.workbook.close();
                } catch (Exception ignored) {}
            }
            throw new IOException("E-UNEXPECTED: Unexpected error opening encrypted workbook: "
                    + e.getMessage(), e);
        }

        return session;
    }

    /**
     * Static method to save a workbook session to its current file path.
     * Delegates to the instance save() method.
     *
     * @param session The WorkbookSession to save
     * @throws IOException if session is null, in read-only mode, or save operation fails
     */
    public static void saveWorkbook(WorkbookSession session) throws IOException {
        if (session == null) {
            throw new IOException("E-SESSION-NULL: WorkbookSession cannot be null.");
        }
        session.save();
    }

    /**
     * Static method to save a workbook session to a new file path.
     * Delegates to the instance saveAs() method.
     *
     * @param session The WorkbookSession to save
     * @param newFilePath The destination file path
     * @param overwrite Whether to overwrite existing files
     * @throws IOException if session is null or save operation fails
     */
    public static void saveWorkbookAs(WorkbookSession session, String newFilePath, boolean overwrite) throws IOException {
        if (session == null) {
            throw new IOException("E-SESSION-NULL: WorkbookSession cannot be null.");
        }
        session.saveAs(newFilePath, overwrite);
    }

    // ========== INSTANCE SAVE METHODS ==========

    /**
     * Saves the current workbook to its current file path.
     * Uses atomic write-to-temp-then-move strategy for Windows reliability.
     *
     * @throws IOException if session is read-only, workbook is null, or I/O error occurs
     */
    public void save() throws IOException {
        if (workbook == null) {
            throw new IOException("E-WB-NULL: No workbook is loaded; nothing to save.");
        }

        if (file == null) {
            throw new IOException("E-PATH-UNSET: Destination path is not set. Use saveAs() to specify a file path.");
        }

        if (readOnly) {
            throw new IOException("E-READONLY: Session is read-only. Use saveAs() to save to a new file or reopen in write mode.");
        }

        validateWorkbookFormat(file);
        saveToFile(file);
    }

    /**
     * Saves the current workbook to a new file path and switches the session to that file.
     *
     * @param newFilePath The destination file path
     * @param overwrite Whether to replace existing files
     * @throws IOException if save operation fails or destination cannot be written
     */
    public void saveAs(String newFilePath, boolean overwrite) throws IOException {
        validateFilePath(newFilePath);

        if (workbook == null) {
            throw new IOException("E-WB-NULL: No workbook is loaded; nothing to save.");
        }

        File newFile = new File(newFilePath);
        Path newPath = newFile.toPath().toAbsolutePath();

        // Create parent directories
        Path parent = newPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        // Check for existing file
        if (newFile.exists() && !overwrite) {
            throw new IOException("E-EXISTS: Destination file already exists: " + newFile.getAbsolutePath()
                    + ". Pass overwrite=true to replace.");
        }

        validateWorkbookFormat(newFile);

        // Release current lock
        releaseLock();

        try {
            // Save to new location
            saveToFile(newFile);

            // Switch session to new file
            this.file = newFile;
            this.readOnly = false;

            // Acquire lock on new file
            acquireLock();

        } catch (IOException e) {
            // Attempt to restore original lock if save failed
            try {
                acquireLock();
            } catch (Exception ignored) {
            }
            throw e;
        }
    }

    // ========== WINDOWS FILE LOCKING (RandomAccessFile) ==========

    /**
     * Acquires a file lock using RandomAccessFile and FileChannel.
     * This is optimized for Windows file locking behavior.
     *
     * - Read-only mode: Shared lock (allows other readers)
     * - Write mode: Exclusive lock (blocks all other access)
     */
    private void acquireLock() throws IOException {
        releaseLock();

        if (file == null) {
            throw new IOException("E-FILE-NULL: File is not set; cannot acquire lock.");
        }

        try {
            // Open RandomAccessFile in appropriate mode
            String mode = readOnly ? "r" : "rw";
            randomAccessFile = new RandomAccessFile(file, mode);
            fileChannel = randomAccessFile.getChannel();

            // Try to acquire lock with timeout
            long startTime = System.currentTimeMillis();

            while (true) {
                try {
                    // Shared lock for read-only, exclusive for write
                    fileLock = fileChannel.tryLock(0L, Long.MAX_VALUE, readOnly);

                    if (fileLock != null) {
                        return; // Lock acquired successfully
                    }
                } catch (OverlappingFileLockException e) {
                    // Another thread in this JVM has a lock - retry
                }

                // Check timeout
                if (System.currentTimeMillis() - startTime > LOCK_TIMEOUT_MS) {
                    throw new IOException("E-LOCK-TIMEOUT: Could not acquire "
                            + (readOnly ? "shared" : "exclusive")
                            + " lock on file within " + LOCK_TIMEOUT_MS + "ms: "
                            + file.getAbsolutePath()
                            + ". The file may be open in another application (e.g., Excel).");
                }

                // Wait before retrying
                try {
                    Thread.sleep(LOCK_RETRY_INTERVAL_MS);
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                    throw new IOException("E-LOCK-INTERRUPTED: Interrupted while waiting for file lock: "
                            + file.getAbsolutePath(), ie);
                }
            }

        } catch (IOException e) {
            // Clean up on failure
            releaseLock();
            throw e;
        }
    }

    /**
     * Releases the current file lock and closes associated resources.
     * Safe to call multiple times.
     */
    private void releaseLock() {
        // Release lock
        if (fileLock != null) {
            try {
                if (fileLock.isValid()) {
                    fileLock.release();
                }
            } catch (Exception ignored) {
            } finally {
                fileLock = null;
            }
        }

        // Close channel
        if (fileChannel != null) {
            try {
                if (fileChannel.isOpen()) {
                    fileChannel.close();
                }
            } catch (Exception ignored) {
            } finally {
                fileChannel = null;
            }
        }

        // Close RandomAccessFile
        if (randomAccessFile != null) {
            try {
                randomAccessFile.close();
            } catch (Exception ignored) {
            } finally {
                randomAccessFile = null;
            }
        }
    }

    // ========== SAVE IMPLEMENTATION ==========

    /**
     * Performs the actual save operation using atomic write strategy.
     * Writes to temporary file first, then moves to target location.
     * This prevents corruption if write is interrupted.
     */
    private void saveToFile(File targetFile) throws IOException {
        File tempFile = null;
        boolean hadLock = (fileLock != null && fileLock.isValid());

        try {
            // Create temporary file in same directory for atomic move
            Path targetPath = targetFile.toPath().toAbsolutePath();
            Path directory = targetPath.getParent();
            String baseName = targetFile.getName();

            tempFile = Files.createTempFile(directory, baseName + "_", TEMP_SUFFIX).toFile();

            // Write workbook to temp file
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                workbook.write(fos);
                fos.flush();
            }

            // Dispose SXSSF temporary files
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }

            // CRITICAL FOR WINDOWS: Release ALL file handles before move
            // 1. Release RandomAccessFile lock
            if (hadLock) {
                releaseLock();
            }

            // 2. For XLSX files, close the underlying OPCPackage
            //    This releases the file handle that WorkbookFactory opened
            if (workbook instanceof XSSFWorkbook) {
                XSSFWorkbook xssf = (XSSFWorkbook) workbook;
                try {
                    // Close the package to release file handle
                    // Don't save - we already wrote to temp file
                    xssf.getPackage().revert();
                } catch (Exception ignored) {
                    // Package might already be closed
                }
            }

            // Small delay to ensure Windows releases file handles
            try {
                Thread.sleep(50);
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
            }

            // Now safe to move temp file to target
            try {
                Files.move(tempFile.toPath(), targetPath,
                        StandardCopyOption.REPLACE_EXISTING,
                        StandardCopyOption.ATOMIC_MOVE);
            } catch (AtomicMoveNotSupportedException e) {
                // Fallback if atomic move not supported
                Files.move(tempFile.toPath(), targetPath,
                        StandardCopyOption.REPLACE_EXISTING);
            }

            tempFile = null; // Successfully moved

            // Reopen the workbook from the saved file
            // This ensures a fresh file handle
            try {
                workbook = WorkbookFactory.create(targetFile);
            } catch (Exception e) {
                throw new IOException("E-REOPEN-FAIL: File saved successfully, but failed to reopen workbook: "
                        + e.getMessage(), e);
            }

        } catch (Exception e) {
            // Clean up temp file on failure
            if (tempFile != null && tempFile.exists()) {
                try {
                    Files.deleteIfExists(tempFile.toPath());
                } catch (Exception ignored) {
                }
            }

            throw new IOException("E-SAVE-FAIL: Failed to save workbook to " + targetFile.getAbsolutePath()
                    + ": " + e.getMessage()
                    + ". Check disk space, permissions, and ensure the file is not open in another application.", e);

        } finally {
            // Re-acquire lock after save
            if (hadLock) {
                try {
                    acquireLock();
                } catch (IOException e) {
                    // Non-fatal - file was saved successfully
                    System.err.println("Warning: File saved but failed to re-acquire lock: " + e.getMessage());
                }
            }
        }
    }

    // ========== VALIDATION METHODS ==========

    /**
     * Validates file path is not null or empty.
     */
    private static void validateFilePath(String filePath) throws IOException {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IOException("E-PATH-UNSET: File path cannot be null or empty.");
        }
    }

    /**
     * Validates that the workbook type matches the file extension.
     * Prevents format mismatches (e.g., saving XSSF as .xls).
     */
    private void validateWorkbookFormat(File targetFile) throws IOException {
        String fileName = targetFile.getName().toLowerCase();

        boolean isXls = fileName.endsWith(".xls");
        boolean isXlsx = fileName.endsWith(".xlsx") || fileName.endsWith(".xlsm");

        if (!isXls && !isXlsx) {
            throw new IOException("E-EXT-UNKNOWN: File must have .xls, .xlsx, or .xlsm extension. Got: " + fileName);
        }

        boolean isHSSF = workbook instanceof HSSFWorkbook;
        boolean isXSSF = workbook instanceof XSSFWorkbook || workbook instanceof SXSSFWorkbook;

        if (isXls && isXSSF) {
            throw new IOException("E-FORMAT-MISMATCH: Cannot save OOXML workbook (XSSF/SXSSF) with .xls extension. Use .xlsx or .xlsm.");
        }

        if (isXlsx && isHSSF) {
            throw new IOException("E-FORMAT-MISMATCH: Cannot save binary workbook (HSSF) with .xlsx extension. Use .xls.");
        }
    }

    // ========== SESSION MANAGEMENT ==========

    /**
     * Closes the workbook and releases all resources.
     * This method is called automatically by Automation Anywhere when the session ends.
     */
    @Override
    public void close() throws IOException {
        try {
            // CRITICAL: Release lock BEFORE closing workbook
            // Otherwise POI cannot save changes when closing
            releaseLock();

            if (workbook != null) {
                // Dispose SXSSF temporary files
                if (workbook instanceof SXSSFWorkbook) {
                    try {
                        ((SXSSFWorkbook) workbook).dispose();
                    } catch (Exception ignored) {
                    }
                }

                workbook.close();
            }
        } finally {
            workbook = null;
            // Ensure lock is released even if workbook.close() fails
            releaseLock();
            closed = true;
        }
    }

    @Override
    public boolean isClosed() {
        return closed;
    }

    // ========== GETTERS ==========

    public Workbook getWorkbook() {
        return workbook;
    }

    public File getFile() {
        return file;
    }

    public String getFilePath() {
        return file != null ? file.getAbsolutePath() : null;
    }

    public boolean isReadOnly() {
        return readOnly;
    }
}
