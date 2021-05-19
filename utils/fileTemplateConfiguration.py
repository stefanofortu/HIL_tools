file_TC_MANUAL_Sheet = [""]
file_TC_MANUAL_Column = ["TEST ID", "PRECONDITIONS", "ACTIONS", "EXPECTED RESULTS"]

file_TC_BUILD_Sheet = [""]
file_TC_BUILD_Column = dict(testID_str="TEST ID",
                            stepDescr_header="STEP DESCRIPTION",
                            precondition_header="PRECONDITIONS",
                            action_header="ACTIONS",
                            expected_header="EXPECTED RESULTS")

file_TC_RUN_Sheet = ["aa", "bb"]
file_TC_RUN_Column = dict(testID_str="TEST ID",
                          stepDescr_str="STEP DESCRIPTION",
                          precondition_str="PRECONDITIONS",
                          action_str="ACTIONS",
                          expected_str="EXPECTED RESULTS")

file_Tools_Init_Sheet = ["aa", "bb"]
file_Tools_Init_Column = ["TEST ID", "ID", "STEP DESCRIPTION", "PRECONDITIONS", "ACTIONS", "EXPECTED RESULTS"]

file_Tools_Actions_Sheet = ["actions"]
file_Tools_Actions_Column = dict(function_header="function",
                                 stepDescr_header="STEP DESCRIPTION",
                                 action_header="PRECONDITIONS/ACTION",
                                 expected_header="EXPECTED RESULTS")

file_Tools_Verify_Sheet = ["actions"]
file_Tools_Verify_Column = dict(function_header="function",
                                stepDescr_header="STEP DESCRIPTION",
                                expected_header="EXPECTED RESULTS")


def format_check_TC_Manual():
    return True


def format_check_TC_Build():
    return True


def format_check_TC_Run():
    return True
