package Handler;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;

import Service.ExportFixtureToExcelService;

public class ExportFixtureCommandHandler extends AbstractHandler {

    @Override
    public Object execute(ExecutionEvent event) {
        ExportFixtureToExcelService.execute();
        return null;
    }
}