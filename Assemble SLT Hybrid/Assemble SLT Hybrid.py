#Author-
#Description-
#FrancisNg
#Assemble an SLT hybrid UAV

import adsk.core
import adsk.fusion
import adsk.cam
import traceback
import pathlib
from . import xlrd
import random

def run(context):
    ui = None
    evaluationsFolder = str(pathlib.Path(__file__).parents[3].joinpath('2---Data')) + '/MBO/evaluations.xls'
    with xlrd.open_workbook(evaluationsFolder, on_demand=True) as book:
        active_rows = book.sheet_by_name('current').col_values(0)
        evaluations = list(range(1, len(active_rows)))
    for evaluation in evaluations:
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface
            doc = app.documents.add(adsk.core.DocumentTypes.FusionDesignDocumentType)
            design = app.activeProduct
            root = design.rootComponent

            joints = root.joints
            geometry = adsk.fusion.JointGeometry
            keypointType = adsk.fusion.JointKeyPointTypes

            importManager = app.importManager
            allOccurrences = root.occurrences
            transform = adsk.core.Matrix3D.create()

    #==============================================================================
    #         Function Definitions
    #==============================================================================

            # Import components from a local location to the current design
            def ImportComponent(numberOfCopies, inputLocation, label):
                occurrenceList = []
                importOptions = importManager.createFusionArchiveImportOptions(inputLocation.as_posix())
                importManager.importToTarget(importOptions, root)
                occurrenceList = [allOccurrences.item(allOccurrences.count-1)]
                occurrenceList[0].component.name = label
                for copy in range(numberOfCopies - 1):
                    allOccurrences.addExistingComponent(occurrenceList[0].component, transform)
                    occurrenceList.append(allOccurrences.item(allOccurrences.count-1))
                return occurrenceList

            # Identify joint location selections
            def selectJointLocation(componentOccurrence, bodyName, index, entityType='face'):
                subcomponents = {'cap': 'Cap (1):1', 'rotor': 'Rotor (1):1', 'base': 'Stator (1):1'}
                occurrence = componentOccurrence
                if bodyName in subcomponents.keys():
                    occurrence = componentOccurrence.childOccurrences.itemByName(subcomponents[bodyName])
                    component = occurrence.component
                else:
                    component = componentOccurrence.component
                selectedBody = component.bRepBodies.itemByName(bodyName)
                if entityType == 'face':
                    selectedFace = selectedBody.faces.item(index)
                    jointLocation = selectedFace.createForAssemblyContext(occurrence)
                    jointGeometry = geometry.createByPlanarFace(jointLocation, None, keypointType.CenterKeyPoint)
                elif entityType == 'edge':
                    selectedEdge = selectedBody.edges.item(index)
                    jointLocation = selectedEdge.createForAssemblyContext(occurrence)
                    jointGeometry = geometry.createByCurve(jointLocation, keypointType.CenterKeyPoint)
                return jointGeometry

            # Create rigid joint
            def createJoint(object_1, object_2, type='rigid', isFlipped='False'):
                jointInput = joints.createInput(object_1, object_2)
                if type=='rigid':
                    jointInput.setAsRigidJointMotion()
                elif type=='revolute':
                    jointInput.setAsRevoluteJointMotion(2) # 2 for z-axis
                    jointInput.angle = adsk.core.ValueInput.createByString(str(random.uniform(0, 360)) + 'deg')
                jointInput.isFlipped = isFlipped
                joint = joints.add(jointInput)
                return joint
    
    #==============================================================================
    #         Import Components
    #==============================================================================

            # Load component selections
            Datafolder = pathlib.Path(__file__).parents[3].joinpath('2---Data')
            plane_selection = 1     # fixed
            beam_selection = 1      # fixed
            with xlrd.open_workbook(str(Datafolder) + '/MBO/evaluations.xls', on_demand=True) as book:
                selection_indices = book.sheet_by_name('current').row_values(evaluation, 0, 2)
                motor_selection, propeller_selection = [int(index) for index in selection_indices]

            # Load component names
            with xlrd.open_workbook(str(Datafolder) + '/MBO/components.xlsx', on_demand=True) as book:
                plane_names = book.sheet_by_name('plane').col_values(0)
                beam_names = book.sheet_by_name('beam').col_values(0)
                motor_names = book.sheet_by_name('motor').col_values(0)
                propeller_names = book.sheet_by_name('propeller').col_values(0)
            plane_name = plane_names[plane_selection]
            beam_name = beam_names[beam_selection]
            motor_name = motor_names[motor_selection]
            propeller_name = propeller_names[propeller_selection]
            
            # Locate storage locations of each component
            CADfolder = pathlib.PureWindowsPath(Datafolder).joinpath('CAD')
            locationPlane = pathlib.PureWindowsPath(CADfolder).joinpath(plane_name + '.f3d')
            locationBeam = pathlib.PureWindowsPath(CADfolder).joinpath(beam_name + '.f3d')
            locationMotor = pathlib.PureWindowsPath(CADfolder).joinpath(motor_name + '.f3d')
            locationPropellerCW = pathlib.PureWindowsPath(CADfolder).joinpath(propeller_name + '-CW.f3d')
            locationPropellerCCW = pathlib.PureWindowsPath(CADfolder).joinpath(propeller_name + '-CCW.f3d')
            
            # Import the components and enable access from lists
            plane = ImportComponent(1, locationPlane, "Plane")
            beam = ImportComponent(2, locationBeam, "Beam")
            motor = ImportComponent(4, locationMotor, "Motor")
            propellerCW = ImportComponent(2, locationPropellerCW, "Propeller (clockwise)")
            propellerCCW = ImportComponent(2, locationPropellerCCW, "Propeller (counter-clockwise)")
            
    #==============================================================================
    #         Assemble
    #==============================================================================

            # Load location indices
            with xlrd.open_workbook(str(Datafolder) + '/MBO/components.xlsx', on_demand=True) as book:
                plane_indices = book.sheet_by_name('joints_plane').row_values(plane_selection, 1, 3)
                left, right = [int(index) for index in plane_indices]

                beam_indices = book.sheet_by_name('joints_beam').row_values(beam_selection, 1, 4)
                front, middle, back = [int(index) for index in beam_indices]

                motor_indices = book.sheet_by_name('joints_motor').row_values(motor_selection, 1, 4)
                cap, rotor, base = [int(index) for index in motor_indices]

                propeller_indices = book.sheet_by_name('joints_propeller').row_values(propeller_selection, 1, 3)
                top, bottom = [int(index) for index in propeller_indices]

            # Identify joint locations
            #
            #                   ^
            #           [M0]   | |   [M2]
            #             |    | |    |
            #       /—————|———————————|—————\
            #      /____[B0]_________[B1]____\
            #             |    | |    | 
            #             |    | |    |
            #           [M1]   | |   [M4]
            #                 _|_|_
            #                /_____\
            #
            #               Top View 

            plane0_left = selectJointLocation(plane[0], 'plane', left, 'face')
            plane0_right = selectJointLocation(plane[0], 'plane', right, 'face')

            beam0_front = selectJointLocation(beam[0], 'beam', front, 'face')
            beam1_front = selectJointLocation(beam[1], 'beam', front, 'face')

            beam0_middle = selectJointLocation(beam[0], 'beam', middle, 'face')
            beam1_middle = selectJointLocation(beam[1], 'beam', middle, 'face')

            beam0_back = selectJointLocation(beam[0], 'beam', back, 'face')
            beam1_back = selectJointLocation(beam[1], 'beam', back, 'face')

            motor0_cap = selectJointLocation(motor[0], 'cap', cap, 'edge')
            motor1_cap = selectJointLocation(motor[1], 'cap', cap, 'edge')
            motor2_cap = selectJointLocation(motor[2], 'cap', cap, 'edge')
            motor3_cap = selectJointLocation(motor[3], 'cap', cap, 'edge')

            motor0_rotor = selectJointLocation(motor[0], 'rotor', rotor, 'edge')
            motor1_rotor = selectJointLocation(motor[1], 'rotor', rotor, 'edge')
            motor2_rotor = selectJointLocation(motor[2], 'rotor', rotor, 'edge')
            motor3_rotor = selectJointLocation(motor[3], 'rotor', rotor, 'edge')

            motor0_base = selectJointLocation(motor[0], 'base', base, 'face')
            motor1_base = selectJointLocation(motor[1], 'base', base, 'face')
            motor2_base = selectJointLocation(motor[2], 'base', base, 'face')
            motor3_base = selectJointLocation(motor[3], 'base', base, 'face')

            propeller0_top = selectJointLocation(propellerCW[0], 'propeller', top, 'edge')
            propeller1_top = selectJointLocation(propellerCCW[0], 'propeller', top, 'edge')
            propeller2_top = selectJointLocation(propellerCCW[1], 'propeller', top, 'edge')
            propeller3_top = selectJointLocation(propellerCW[1], 'propeller', top, 'edge')

            propeller0_bottom = selectJointLocation(propellerCW[0], 'propeller', bottom, 'edge')
            propeller1_bottom = selectJointLocation(propellerCCW[0], 'propeller', bottom, 'edge')
            propeller2_bottom = selectJointLocation(propellerCCW[1], 'propeller', bottom, 'edge')
            propeller3_bottom = selectJointLocation(propellerCW[1], 'propeller', bottom, 'edge')

            # Create joints
            plane[0].isGrounded = True

            PB0 = createJoint(beam0_middle, plane0_left, type='rigid', isFlipped=True)
            PB1 = createJoint(beam1_middle, plane0_right, type='rigid', isFlipped=True)

            BMb0 = createJoint(motor0_base, beam0_front, type='revolute', isFlipped=True)
            BMb1 = createJoint(motor1_base, beam0_back, type='revolute', isFlipped=True)
            BMb2 = createJoint(motor2_base, beam1_front, type='revolute', isFlipped=True)
            BMb3 = createJoint(motor3_base, beam1_back, type='revolute', isFlipped=True)
            
            MrP0 = createJoint(propeller0_top, motor0_rotor, type='revolute', isFlipped=False)
            MrP1 = createJoint(propeller1_top, motor1_rotor, type='revolute', isFlipped=True)
            MrP2 = createJoint(propeller2_top, motor2_rotor, type='revolute', isFlipped=True)
            MrP3 = createJoint(propeller3_top, motor3_rotor, type='revolute', isFlipped=False)

            PMc0 = createJoint(motor0_cap, propeller0_bottom, type='rigid', isFlipped=top<bottom) 
            PMc1 = createJoint(motor1_cap, propeller1_bottom, type='rigid', isFlipped=top>bottom) 
            PMc2 = createJoint(motor2_cap, propeller2_bottom, type='rigid', isFlipped=top>bottom)
            PMc3 = createJoint(motor3_cap, propeller3_bottom, type='rigid', isFlipped=top<bottom)
                #The boolean based on empirical observation of the influence of index order on the assembly.
                #The difference in order is probably due to an inconsistency in the modeling procedure.

            # Hide joints as they are already properly placed
            for joint in [PB0, PB1, BMb0, BMb1, BMb2, BMb3, MrP0, MrP1, MrP2, MrP3, PMc0, PMc1, PMc2, PMc3]:
                joint.isLightBulbOn = False

    #==============================================================================
    #         Update Parameters
    #==============================================================================

            with xlrd.open_workbook(str(Datafolder) + '/MBO/evaluations.xls', on_demand=True) as book:
                parameters = book.sheet_by_name('current').row_values(evaluation, 4, 7)
                distanceFromCenterline, beam_length, pitch = [str(parameter) for parameter in parameters]

            params = design.allParameters
            params.itemByName('distanceFromCenterline').expression = distanceFromCenterline + 'cm'
            params.itemByName('beamLength').expression = beam_length + 'cm'
            params.itemByName('pitch').expression = str(float(pitch)+90) + 'deg'       
                # after the 28Oct2019 update, importing somehow turns every import by 90 degrees
            params.itemByName('connectionWidth').expression = params.itemByName('beamWidth').expression

    #==============================================================================
    #         Combine Bodies
    #==============================================================================

            taget_body = plane[0].component.bRepBodies.itemByName('plane')
            tool_bodies = adsk.core.ObjectCollection.create()
            for occurrence in beam + motor + propellerCW + propellerCCW:
                for body in occurrence.bRepBodies:
                    tool_bodies.add(body)
                for subcomponent in occurrence.childOccurrences:
                    for body in subcomponent.bRepBodies:
                        tool_bodies.add(body)          
            combineBodies_input = root.features.combineFeatures.createInput(taget_body, tool_bodies)
            combineBodies = root.features.combineFeatures.add(combineBodies_input)
            for occurrence in beam + motor + propellerCW + propellerCCW:
                removeBodies = root.features.removeFeatures.add(occurrence)

    #==============================================================================
    #         Export STEP Format
    #==============================================================================

            CFDfolder = pathlib.PureWindowsPath(Datafolder).joinpath('CFD')
            filename = ('M' + str(motor_selection) + "_"
                        + 'P' + str(propeller_selection) + "_"
                        + distanceFromCenterline + 'cm_' 
                        + beam_length + 'cm_' 
                        + pitch + 'deg')
            filename = filename.replace('.', ',') + '.step'

            export_directory = pathlib.PureWindowsPath(CFDfolder).joinpath(filename)
            STEP_Options = design.exportManager.createSTEPExportOptions(str(export_directory))
            STEP_Export = design.exportManager.execute(STEP_Options)

    #==============================================================================
    #         Error Message
    #==============================================================================

        except:
            print('Failed:\n{}'.format(traceback.format_exc()))
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
