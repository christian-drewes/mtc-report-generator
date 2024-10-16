/*
 * Copyright (c) 2023. PortSwigger Ltd. All rights reserved.
 *
 * This code may be used to extend the functionality of Burp Suite Community Edition
 * and Burp Suite Professional, provided that this usage does not violate the
 * license terms for those products.
 */

package example.contextmenu;

import burp.api.montoya.MontoyaApi;
import burp.api.montoya.core.ToolType;
import burp.api.montoya.http.message.HttpRequestResponse;
import burp.api.montoya.ui.contextmenu.ContextMenuEvent;
import burp.api.montoya.ui.contextmenu.ContextMenuItemsProvider;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.*;
import java.awt.*;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.io.File;
import java.io.FileOutputStream;

public class MyContextMenuItemsProvider implements ContextMenuItemsProvider
{

    private final MontoyaApi api;

    public MyContextMenuItemsProvider(MontoyaApi api)
    {
        this.api = api;
    }

    @Override
    public List<Component> provideMenuItems(ContextMenuEvent event)
    {
        if (event.isFromTool(ToolType.PROXY, ToolType.TARGET, ToolType.LOGGER))
        {
            List<Component> menuItemList = new ArrayList<>();

            JMenuItem retrieveRequestItem = new JMenuItem("Print request");
            JMenuItem retrieveResponseItem = new JMenuItem("Print response");
            JMenuItem captureMTC = new JMenuItem("Print MTC1A-TEST4");

            HttpRequestResponse requestResponse = event.messageEditorRequestResponse().isPresent() ? event.messageEditorRequestResponse().get().requestResponse() : event.selectedRequestResponses().get(0);

            retrieveRequestItem.addActionListener(l -> api.logging().logToOutput("Request is:\r\n" + requestResponse.request().toString()));
            menuItemList.add(retrieveRequestItem);
            menuItemList.add(captureMTC);

            String endpoint = requestResponse.request().path();
            Short statusCode = requestResponse.response().statusCode();

            if (requestResponse.response() != null)
            {
                retrieveResponseItem.addActionListener(l -> api.logging().logToOutput("Response is:\r\n" + requestResponse.response().toString()));
                menuItemList.add(retrieveResponseItem);
            }

            // MTC Request/Response capture
            captureMTC.addActionListener(l -> {

                api.logging().logToOutput("Request is:\r\n" + requestResponse.request().toString());

                try {
                    api.logging().logToOutput("BEFORE");
                    XWPFDocument document = new XWPFDocument();

                    //Write the Document in file system
                    FileOutputStream out = new FileOutputStream( new File("C:/tmp/createdocument.docx"));

                    //create Paragraph
                    XWPFRun run = newLine(document, true);
                    run.setText(endpoint + " API call:\r\n");
                    run = newLine(document, false);
                    run.setText(requestResponse.request().toString() + "\r\n");

                    if (requestResponse.response() != null && statusCode <= 301)
                    {
                        run = newLine(document, true);
                        run.setText("Correct allow responses:");
                        run = newLine(document, false);
                        run.setText(requestResponse.response().toString());
                    } else if (requestResponse.response() != null && statusCode > 301) {
                        run = newLine(document, true);
                        run.setText("Incorrect DENY response:");
                        run = newLine(document, false);
                        run.setText(requestResponse.response().toString());
                    } else {
                        api.logging().logToOutput("Response is:\r\n No Response");
                        run.setText("No Response");
                    }

                    document.write(out);
                    out.close();

                } catch (Exception e) {
                    api.logging().logToOutput("Don't freak out: " + e);
                }
            });

            return menuItemList;
        }

        return null;
    }

    XWPFRun newLine(XWPFDocument document, Boolean isBold){
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(9);
        run.setBold(isBold);

        return run;
    }

    XWPFRun newLine(XWPFDocument document){
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(9);
        run.setBold(false);

        return run;
    }
}