package cn.afterturn.easypoi.pdf.entity;

import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.pdf.handler.IPdfCellHandler;
import cn.afterturn.easypoi.pdf.styler.IPdfExportStyler;
import lombok.Data;

@Data
public class PdfImportParams extends ImportParams {
    
    private IPdfCellHandler cellHandler;
}
