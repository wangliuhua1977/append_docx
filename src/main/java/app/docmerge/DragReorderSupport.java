package app.docmerge;

import javax.swing.JComponent;
import javax.swing.JTable;
import javax.swing.TransferHandler;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;

public class DragReorderSupport extends TransferHandler {
    private static final DataFlavor FLAVOR = new DataFlavor(int[].class, "行索引");
    private final JTable table;
    private final ReorderCallback callback;
    private int[] rows;

    public DragReorderSupport(JTable table, ReorderCallback callback) {
        this.table = table;
        this.callback = callback;
    }

    @Override
    protected Transferable createTransferable(JComponent c) {
        rows = table.getSelectedRows();
        return new RowsTransferable(rows);
    }

    @Override
    public int getSourceActions(JComponent c) {
        return MOVE;
    }

    @Override
    public boolean canImport(TransferSupport support) {
        return support.isDataFlavorSupported(FLAVOR);
    }

    @Override
    public boolean importData(TransferSupport support) {
        if (!canImport(support)) {
            return false;
        }
        try {
            int[] from = (int[]) support.getTransferable().getTransferData(FLAVOR);
            JTable.DropLocation dl = (JTable.DropLocation) support.getDropLocation();
            int index = dl.getRow();
            if (index < 0) {
                index = table.getRowCount();
            }
            callback.onReorder(from, index);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    @Override
    protected void exportDone(JComponent source, Transferable data, int action) {
        rows = null;
    }

    public interface ReorderCallback {
        void onReorder(int[] fromRows, int toRow);
    }

    private static class RowsTransferable implements Transferable {
        private final int[] rows;

        private RowsTransferable(int[] rows) {
            this.rows = rows;
        }

        @Override
        public DataFlavor[] getTransferDataFlavors() {
            return new DataFlavor[]{FLAVOR};
        }

        @Override
        public boolean isDataFlavorSupported(DataFlavor flavor) {
            return FLAVOR.equals(flavor);
        }

        @Override
        public Object getTransferData(DataFlavor flavor) {
            if (!isDataFlavorSupported(flavor)) {
                return null;
            }
            return rows;
        }
    }
}
