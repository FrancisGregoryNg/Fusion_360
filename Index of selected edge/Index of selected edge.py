#Author-
#Description-https://forums.autodesk.com/t5/fusion-360-api-and-scripts/how-do-you-know-item-id-of-selected-faces/td-p/7517110

import adsk.core, adsk.fusion, traceback

def run(context):
    ui = None
    try: 
        app = adsk.core.Application.get()
        ui = app.userInterface
     
        edgeSel = ui.selectEntity('Select an edge.', 'Edges')
        if edgeSel:
            selectedEdge = adsk.fusion.BRepEdge.cast(edgeSel.entity)
            selectedBody = selectedEdge.body
            selectedComponent = selectedBody.parentComponent
            
            # Find the index of this face within the body.            
            edgeIndex = -1
            edgeCount = 0
            for edge in selectedBody.edges:
                if edge == selectedEdge:
                    edgeIndex = edgeCount
                    break
                edgeCount += 1
                
            print('Edge #' + str(edgeIndex))
            ui.messageBox('Edge #' + str(edgeIndex))                
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
