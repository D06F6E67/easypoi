package cn.afterturn.easypoi.pdf.watermark;

import lombok.Data;
import org.apache.pdfbox.pdmodel.PDPage;

import java.util.List;

@Data
public class RemoveResult {
    PDPage page;
    int pageNo;
    List<?> tokens;

    public RemoveResult(PDPage page, int pageNo, List<?> tokens) {
        this.page = page;
        this.pageNo = pageNo;
        this.tokens = tokens;
    }

}
