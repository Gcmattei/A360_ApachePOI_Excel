package com.davita.botcommand.excel.sessions;

import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.channels.OverlappingFileLockException;
import java.nio.channels.FileChannel;
import java.nio.channels.FileLock;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;

/**
 * Session for holding a POI Workbook and managing cross-platform advisory file locks via Java NIO.
 */
public class WorkbookSession implements CloseableSessionObject {

    private static final long DEFAULT_LOCK_TIMEOUT_MS = 10_000L; // 10 seconds
    private static final long LOCK_RETRY_SLEEP_MS = 100L;

    private volatile boolean closed = false;
    private Workbook workbook;
    private String filePath;
    private boolean readOnly;

    // NIO locking state
    private FileChannel channel;
    private FileLock nioLock;

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public boolean isReadOnly() {
        return readOnly;
    }

    public void setReadOnly(boolean readOnly) {
        this.readOnly = readOnly;
    }

    @Override
    public boolean isClosed() {
        return closed;
    }

    /**
     * Acquire an advisory lock on the current filePath using Java NIO.
     * Uses tryLock with retries up to a timeout to avoid indefinite blocking.
     */
    public void acquireLock() throws IOException {
        releaseLock();

        if (filePath == null || filePath.trim().isEmpty()) {
            return;
        }

        final Path path = Path.of(filePath);
        final String os = System.getProperty("os.name", "unknown");
        String fs = "unknown";
        try { fs = Files.getFileStore(path).name(); } catch (Exception ignore) {}

        if (readOnly) {
            if (!Files.exists(path)) {
                throw new IOException("E-LOCK-NOTFOUND: Cannot acquire read lock because file does not exist: "
                        + path + " (OS=" + os + ", FS=" + fs + ").");
            }
            channel = FileChannel.open(path, StandardOpenOption.READ);
            nioLock = tryLockWithTimeout(channel, true, DEFAULT_LOCK_TIMEOUT_MS, os, fs);
        } else {
            // For writable sessions, create the file if it does not exist to allow locking
            channel = FileChannel.open(path,
                    StandardOpenOption.READ,
                    StandardOpenOption.WRITE,
                    StandardOpenOption.CREATE);
            nioLock = tryLockWithTimeout(channel, false, DEFAULT_LOCK_TIMEOUT_MS, os, fs);
        }
    }

    private FileLock tryLockWithTimeout(FileChannel ch,
                                        boolean shared,
                                        long timeoutMs,
                                        String os,
                                        String fs) throws IOException {
        final long deadline = System.currentTimeMillis() + timeoutMs;
        while (true) {
            try {
                FileLock fl = ch.tryLock(0, Long.MAX_VALUE, shared);
                if (fl != null) {
                    return fl;
                }
            } catch (OverlappingFileLockException e) {
                // Another lock in this JVM overlaps; retry until timeout
            } catch (IOException ioe) {
                // Bubble up with context for unsupported filesystems or other I/O errors
                throw new IOException("E-LOCK-IO: Failed to acquire " + (shared ? "shared" : "exclusive")
                        + " lock on " + filePath + " (OS=" + os + ", FS=" + fs + "): "
                        + ioe.getMessage(), ioe);
            }

            if (System.currentTimeMillis() >= deadline) {
                throw new IOException("E-LOCK-TIMEOUT: Could not acquire "
                        + (shared ? "shared" : "exclusive") + " lock within "
                        + timeoutMs + "ms on " + filePath + " (OS=" + os + ", FS=" + fs
                        + "); ensure the file is not open elsewhere and that the filesystem supports locking.");
            }

            try {
                Thread.sleep(LOCK_RETRY_SLEEP_MS);
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
                throw new IOException("E-LOCK-INTERRUPTED: Interrupted while waiting for file lock on "
                        + filePath + ".", ie);
            }
        }
    }

    /**
     * Release any held advisory lock and close the channel.
     */
    public void releaseLock() {
        // Release NIO lock if present
        try {
            if (nioLock != null && nioLock.isValid()) {
                nioLock.release();
            }
        } catch (Exception ignored) {
        } finally {
            nioLock = null;
        }

        // Close channel if present
        try {
            if (channel != null && channel.isOpen()) {
                channel.close();
            }
        } catch (Exception ignored) {
        } finally {
            channel = null;
        }
    }

    /**
     * Switch the lock to a new file path, preserving the current readOnly policy.
     * If acquiring on the new path fails, attempts to restore the previous lock.
     */
    public void switchLockTo(String newPath) throws IOException {
        final String oldPath = this.filePath;
        releaseLock();
        this.filePath = newPath;
        try {
            acquireLock();
        } catch (IOException first) {
            try {
                this.filePath = oldPath;
                acquireLock();
            } catch (Exception ignored) {
                this.filePath = oldPath;
            }
            throw first;
        }
    }

    @Override
    public void close() throws IOException {
        try {
            if (workbook != null) {
                // Best-effort cleanup for SXSSF temporary files
                if (workbook instanceof SXSSFWorkbook) {
                    try {
                        ((SXSSFWorkbook) workbook).dispose();
                    } catch (Exception ignore) {
                        // swallow; disposal is best-effort on close
                    }
                }
                workbook.close();
            }
        } finally {
            workbook = null;
            releaseLock();
            closed = true;
        }
    }

    public void saveChanges() throws IOException {
        if (workbook == null) {
            throw new IOException("E-WB-NULL: No workbook is loaded; nothing to save.");
        }

        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IOException("E-PATH-UNSET: Destination path is not set; call saveAs(...) to choose a file path.");
        }

        if (readOnly) {
            throw new IOException("E-READONLY: Session is read-only; use saveAs(...) to a writable path or reopen without readOnly.");
        }

        final Path target = Path.of(filePath);
        final Path dir = target.toAbsolutePath().getParent();
        if (dir == null) {
            throw new IOException("E-DIR-RESOLVE: Cannot resolve parent directory for: " + target + ".");
        }
        if (Files.exists(target) && Files.isDirectory(target)) {
            throw new IOException("E-TARGET-IS-DIR: Destination is a directory, not a file: " + target + ".");
        }

        // Validate workbook type vs extension (HSSFWorkbook -> .xls; XSSFWorkbook/SXSSFWorkbook -> .xlsx/.xlsm)
        final String name = target.getFileName().toString();
        final String ext = name.contains(".") ? name.substring(name.lastIndexOf('.') + 1).toLowerCase() : "";
        final boolean isXls = ext.equals("xls");
        final boolean isXlsxLike = ext.equals("xlsx") || ext.equals("xlsm") || ext.equals("xltx") || ext.equals("xltm");
        final boolean isHssf = workbook instanceof org.apache.poi.hssf.usermodel.HSSFWorkbook;
        final boolean isXssf = workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook;
        final boolean isSxssf = workbook instanceof org.apache.poi.xssf.streaming.SXSSFWorkbook;

        if (isXls && (isXssf || isSxssf)) {
            throw new IOException("E-FORMAT-MISMATCH: OOXML workbook cannot be saved with .xls; use .xlsx or .xlsm via saveAs(...).");
        }
        if (isXlsxLike && isHssf) {
            throw new IOException("E-FORMAT-MISMATCH: Binary .xls workbook cannot be saved as .xlsx; use .xls via saveAs(...).");
        }
        if (!isXls && !isXlsxLike) {
            throw new IOException("E-EXT-UNKNOWN: Unsupported or missing extension for Excel file: " + name
                    + " (expected .xls or .xlsx/.xlsm).");
        }

        final boolean hadLock = (nioLock != null && nioLock.isValid()) || (channel != null && channel.isOpen());
        Path temp = null;

        try {
            if (hadLock) {
                releaseLock();
            }

            Files.createDirectories(dir);
            temp = Files.createTempFile(dir, name + ".", ".tmp_" + System.nanoTime());

            try (OutputStream os = new BufferedOutputStream(
                    Files.newOutputStream(temp,
                            java.nio.file.StandardOpenOption.WRITE,
                            java.nio.file.StandardOpenOption.TRUNCATE_EXISTING))) {
                workbook.write(os);    // Persist workbook data
                os.flush();
            }

            // Remove POI's streaming temp files for SXSSF workbooks
            if (isSxssf) {
                ((org.apache.poi.xssf.streaming.SXSSFWorkbook) workbook).dispose();
            }

            try {
                Files.move(temp, target,
                        StandardCopyOption.REPLACE_EXISTING,
                        StandardCopyOption.ATOMIC_MOVE);
            } catch (java.nio.file.AtomicMoveNotSupportedException e) {
                // Fallback when atomic moves are not supported (e.g., cross-volume)
                Files.move(temp, target, StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException ioe) {
            // Ensure temporary file is deleted on any failure to avoid littering the filesystem
            safeDelete(temp);
            throw new IOException("E-SAVE-FAIL: Failed saving to " + filePath + " ("
                    + ioe.getClass().getSimpleName() + "): " + ioe.getMessage()
                    + ". Check permissions, disk space, antivirus/file-sync, and that the file isn’t open elsewhere.", ioe);
        } finally {
            // Defensive cleanup if temp still exists (e.g., partial failure scenarios)
            safeDelete(temp);

            if (hadLock) {
                try {
                    acquireLock();
                } catch (IOException re) {
                    throw new IOException("E-RELOCK-FAIL: Saved file, but failed to re-acquire session lock: "
                            + re.getMessage() + ". The file was saved; reopen the session if locking is required.", re);
                }
            }
        }
    }

    /** Delete a file if it exists, swallowing any exception (best-effort). */
    private static void safeDelete(Path p) {
        if (p == null) return;
        try {
            java.nio.file.Files.deleteIfExists(p);
        } catch (Exception ignored) {
            // best-effort cleanup; nothing else to do
        }
    }
    /**
     * Convenience: Save the workbook to a new path, optionally allowing overwrite.
     */
    public void saveAs(String newPath, boolean overwrite) throws IOException {
        if (newPath == null || newPath.trim().isEmpty()) {
            throw new IOException("E-PATH-UNSET: Destination path is not set; provide a non-empty path.");
        }
        Path newTarget = Path.of(newPath);
        if (Files.exists(newTarget) && !overwrite && Files.isRegularFile(newTarget)) {
            throw new IOException("E-EXISTS: Destination already exists: " + newTarget + "; pass overwrite=true to replace.");
        }

        switchLockTo(newPath);
        saveChanges();
    }
}
