package pl.jsolve.templ4docx.executor;

import java.util.List;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import pl.jsolve.sweetener.collection.Collections;
import pl.jsolve.sweetener.text.Strings;
import pl.jsolve.templ4docx.condition.ConditionSplitter;
import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.insert.ConditionInsert;
import pl.jsolve.templ4docx.util.Condition;
import pl.jsolve.templ4docx.variable.Variables;

public class ConditionExecutor {

    private final static String CONDITION_PREFIX = "<docx:if";
    private final static String CONDITION_SUFFIX = "</docx:if>";

    private Variables variables;
    private ParagraphRemover paragraphRemover;
    private ConditionComparator conditionComparator;
    private ConditionSplitter conditionSplitter;

    public ConditionExecutor(Variables variables) {
        this.variables = variables;
        this.paragraphRemover = new ParagraphRemover();
        this.conditionComparator = new ConditionComparator();
        this.conditionSplitter = new ConditionSplitter();
    }

    public void execute(Docx docx) {
        XWPFDocument xwpfDocument = docx.getXWPFDocument();
        execute(xwpfDocument);

        executeForTables(xwpfDocument.getTables());
    }

    // IBody
    private void executeForTables(List<XWPFTable> tables) {
        for (XWPFTable tbl : tables) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    if (!cell.getTables().isEmpty()) {
                        executeForTables(cell.getTables());
                    }
                    execute(cell);
                }
            }
        }
    }

    private void execute(IBody ibody) {
        while (true) {
            ConditionInsert conditionInsert = findConditionInsert(ibody.getBodyElements());
            if (!conditionInsert.isFound() || conditionInsert.getStartIndex() == null
                    || conditionInsert.getEndIndex() == null) {
                break;
            }
            executeCondition(ibody, conditionInsert);
        }
    }

    private ConditionInsert findConditionInsert(List<IBodyElement> bodyElements) {

        ConditionInsert conditionInsert = new ConditionInsert();

        int deepIndex = 0;
        boolean found = false;
        for (int i = 0; i < bodyElements.size(); i++) {

            if (bodyElements.get(i).getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph paragraph = (XWPFParagraph) bodyElements.get(i);
                String paragraphText = paragraph.getText();
                if (paragraphText.contains(CONDITION_PREFIX)) {
                    if (!found) {
                        found = true;
                        conditionInsert.setStartIndex(i);
                        conditionInsert.setStartParagraph(paragraph);
                        int startIndex = paragraphText.indexOf(CONDITION_PREFIX);
                        List<Integer> indexesOf = Strings.indexesOf(paragraphText, ">");
                        int closeTagIndex = -1;
                        for (Integer index : indexesOf) {
                            if (index < startIndex) {
                                continue;
                            }
                            if (index > 0 && paragraphText.charAt(index - 1) != '/') {
                                closeTagIndex = index;
                                break;
                            }
                        }
                        conditionInsert.setCondition(paragraphText.substring(startIndex, closeTagIndex + 1));
                        deepIndex++;
                    } else {
                        deepIndex++;
                    }
                    if (paragraphText.contains(CONDITION_SUFFIX)) {
                        deepIndex--;
                        conditionInsert.setEndIndex(i);
                        conditionInsert.setEndParagraph(paragraph);
                    }
                }
                if (paragraphText.contains(CONDITION_SUFFIX)) {
                    if (deepIndex == 1) {
                        conditionInsert.setEndIndex(i);
                        conditionInsert.setEndParagraph(paragraph);
                        conditionInsert.setFound(true);
                        deepIndex = 0;
                    } else {
                        deepIndex--;
                    }
                }
            }
        }
        conditionInsert.setFound(found);
        return conditionInsert;
    }

    private void executeCondition(IBody xwpfDocument, ConditionInsert conditionInsert) {
        Integer startIndex = conditionInsert.getStartIndex();
        Integer endIndex = conditionInsert.getEndIndex();

        List<IBodyElement> documentBodyElements = xwpfDocument.getBodyElements();
        List<IBodyElement> listOfBodyElements = Collections.newArrayList();
        for (int i = startIndex; i <= endIndex; i++) {
        	listOfBodyElements.add(documentBodyElements.get(i));
        }
        IBodyElement[] bodyElements = listOfBodyElements.toArray(new IBodyElement[0]);

        Condition condition = conditionSplitter.splitCondition(conditionInsert.getCondition());
        boolean isConditionFulfilled = conditionComparator.compare(condition, variables);

        paragraphRemover.removeConditionTagsFromParagraphs(conditionInsert, xwpfDocument, isConditionFulfilled,
                bodyElements);
    }

}
