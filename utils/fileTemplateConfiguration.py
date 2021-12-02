file_TC_MANUAL_Sheet = [""]
file_TC_MANUAL_Column = ["TEST ID", "PRECONDITIONS", "ACTIONS", "EXPECTED RESULTS"]

file_TC_BUILD_Sheet = [""]
file_TC_BUILD_Column = dict(enable_header="ENABLE",
                            testN_header="TEST N.",
                            testID_header="TEST ID",
                            stepDescr_header="STEP DESCRIPTION",
                            testType_header="TC Type",
                            stepID_header="ID",
                            precondition_header="PRECONDITIONS / OPERATION",
                            action_header="ACTIONS",
                            expected_header="EXPECTED RESULT",
                            timeStep_header="Time step [ms]",
                            sampleTime_header="Sample time [ms]",
                            tolerance_header="Tolerance [%]")

file_TC_RUN_Sheet = ["aa", "bb"]
file_TC_RUN_Column = dict(enable_header="ENABLE",
                          testN_header="TEST N.",
                          testID_header="TEST ID",
                          stepID_header="ID",
                          stepDescr_header="STEP DESCRIPTION",
                          precondition_header="PRECONDITIONS / OPERATION",
                          action_header="ACTIONS",
                          expected_header="EXPECTED RESULT",
                          timeStep_header="Time step [ms]",
                          sampleTime_header="Sample time [ms]",
                          tolerance_header="Tolerance [%]")

file_Tools_Init_Sheet = ["aa", "bb"]
file_Tools_Init_Column = ["TEST ID", "ID", "STEP DESCRIPTION", "PRECONDITIONS", "ACTIONS", "EXPECTED RESULTS"]

file_Tools_Actions_Sheet = ["actions"]
file_Tools_Actions_Column = dict(function_header="function",
                                 stepDescr_header="STEP DESCRIPTION",
                                 action_header="PRECONDITIONS/ACTION",
                                 expected_header="EXPECTED RESULTS",
                                 timeStep_header="Time step [ms]",
                                 sampleTime_header="Sample time [ms]",
                                 tolerance_header="Tolerance [%]")

file_Tools_Verify_Sheet = ["actions"]
file_Tools_Verify_Column = dict(fieldType_header="fieldType",
                                value_header="value",
                                enable_header="ENABLE",
                                stepDescr_header="STEP DESCRIPTION",
                                precondition_header="PRECONDITIONS/ACTION",
                                expected_header="EXPECTED RESULTS",
                                timeStep_header="Time step [ms]",
                                sampleTime_header="Sample time [ms]",
                                tolerance_header="Tolerance [%]")
