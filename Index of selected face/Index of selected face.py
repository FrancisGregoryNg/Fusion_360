#Author-
#Description-https://forums.autodesk.com/t5/fusion-360-api-and-scripts/how-do-you-know-item-id-of-selected-faces/td-p/7517110

import adsk.core, adsk.fusion, traceback

def run(context):
    ui = None
    try: 
        app = adsk.core.Application.get()
        ui = app.userInterface
     
        faceSel = ui.selectEntity('Select a face.', 'Faces')
        if faceSel:
            selectedFace = adsk.fusion.BRepFace.cast(faceSel.entity)
            selectedBody = selectedFace.body
            selectedComponent = selectedBody.parentComponent
            
            # Find the index of this face within the body.            
            faceIndex = -1
            faceCount = 0
            for face in selectedBody.faces:
                if face == selectedFace:
                    faceIndex = faceCount
                    break
                faceCount += 1
                
            print('Face #' + str(faceIndex))
            ui.messageBox('Face #' + str(faceIndex))                
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
