package com.example.docmerger.ui;

import com.example.docmerger.model.FileTableModel;

import javax.swing.JComponent;
import javax.swing.JTable;
import javax.swing.TransferHandler;
import javax.swing.table.TableModel;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.util.Arrays;
import java.util.concurrent.atomic.AtomicBoolean;

public final class ReorderableTableTransferHandler extends TransferHandler {
    private static final DataFlavor ROW_FLAVOR = new DataFlavor(int[].class, "RowIndices");

    private final JTable table;
    private final AtomicBoolean dragEnabled;
    private final RowMoveListener moveListener;

    public ReorderableTableTransferHandler(JTable table, AtomicBoolean dragEnabled, RowMoveListener moveListener) {
        this.table = table;
        this.dragEnabled = dragEnabled;
        this.moveListener = moveListener;
    }

    @Override
    protected Transferable createTransferable(JComponent c) {
        int[] rows = table.getSelectedRows();
        return new RowTransferable(rows);
    }

    @Override
    public int getSourceActions(JComponent c) {
        return MOVE;
    }

    @Override
    public boolean canImport(TransferSupport support) {
        if (!dragEnabled.get()) {
            return false;
        }
        if (!support.isDrop()) {
            return false;
        }
        if (support.getComponent() != table) {
            return false;
        }
        return support.isDataFlavorSupported(ROW_FLAVOR);
    }

    @Override
    public boolean importData(TransferSupport support) {
        if (!canImport(support)) {
            return false;
        }
        JTable.DropLocation dropLocation = (JTable.DropLocation) support.getDropLocation();
        int dropIndex = dropLocation.getRow();
        if (dropIndex < 0) {
            dropIndex = table.getRowCount();
        }
        try {
            int[] rows = (int[]) support.getTransferable().getTransferData(ROW_FLAVOR);
            if (rows.length == 0) {
                return false;
            }
            Arrays.sort(rows);
            TableModel model = table.getModel();
            if (model instanceof FileTableModel fileTableModel) {
                fileTableModel.moveRows(rows, dropIndex);
                moveListener.rowsMoved(rows, dropIndex);
                return true;
            }
        } catch (UnsupportedFlavorException | IOException ignored) {
            return false;
        }
        return false;
    }

    private static final class RowTransferable implements Transferable {
        private final int[] rows;

        private RowTransferable(int[] rows) {
            this.rows = rows == null ? new int[0] : rows.clone();
        }

        @Override
        public DataFlavor[] getTransferDataFlavors() {
            return new DataFlavor[]{ROW_FLAVOR};
        }

        @Override
        public boolean isDataFlavorSupported(DataFlavor flavor) {
            return ROW_FLAVOR.equals(flavor);
        }

        @Override
        public Object getTransferData(DataFlavor flavor) throws UnsupportedFlavorException {
            if (!isDataFlavorSupported(flavor)) {
                throw new UnsupportedFlavorException(flavor);
            }
            return rows.clone();
        }
    }

    @FunctionalInterface
    public interface RowMoveListener {
        void rowsMoved(int[] rows, int targetIndex);
    }
}
