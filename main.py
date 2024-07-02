"""This module contains the business logic of the function.

Use the automation_context module to wrap your function in an Autamate context helper
"""

from pydantic import Field, SecretStr
from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)

from flatten import flatten_base
from specklepy import objects
from specklepy.objects.geometry import Point #Import point to search for points in the flattened base
import pandas as pd

class FunctionInputs(AutomateBase):
    """These are function author defined values.

    Automate will make sure to supply them matching the types specified here.
    Please use the pydantic model schema to define your inputs:
    https://docs.pydantic.dev/latest/usage/models/
    """

    # an example how to use secret values
    whisper_message: SecretStr = Field(title="Write something here to test the function inputs")
    # search_speckle_type: str = Field(
    #     title="Speckle type to search the model for",
    #     description=(
    #         "If a object has the following speckle_type,"
    #         " it will be counted."
    #     ),
    # )


def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    """This is a Speckle Automate function to extract the model's Points and their XY coordinates. 

    Args:
        automate_context: A context helper object, that carries relevant information
            about the runtime context of this function.
            It gives access to the Speckle project data, that triggered this run.
            It also has conveniece methods attach result data to the Speckle model.
        function_inputs: An instance object matching the defined schema.
    """
    # the context provides a conveniet way, to receive the triggering version
    version_root_object = automate_context.receive_version()

    objects_with_search_speckle_type = [
        b
        for b in flatten_base(version_root_object)
        if b.speckle_type == Point
    ]
    count = len(objects_with_search_speckle_type)


    if count > 0:
        # this is how a run is marked with a failure cause
        automate_context.attach_result_to_objects(
            category="Point Types",
            # object_ids=[o.id for o in objects_with_search_speckle_type if o.id],
            message=f"This model has {count} Point data objects present."
        )

        df = pd.DataFrame()

        df['Point_Number'] = None
        df['X-Coordinate'] = None
        df['Y-Coordinate'] = None
        
        for i in range(0, count):
            point = objects_with_search_speckle_type[i]
            df.loc[i, 'Point_Number'] = i
            df.loc[i, 'X-Coordinate'] = point.x
            df.loc[i, 'Y-Coordinate'] = point.y

        excel_file_path = 'Model_Coordinates.xlsx'
        df.to_excel(excel_file_path, index=False)
        automate_context.store_file_result("./Model_Coordinates.xlsx")

        automate_context.mark_run_success(
            "Automation successfully run: "
            f"Found {count} Points in the model"
        )

        # set the automation context view, to the original model / version view
        # to show the offending objects
        automate_context.set_context_view()

    else:
        automate_context.mark_run_success("No Point types found in the model.")

    # if the function generates file results, this is how it can be
    # attached to the Speckle project / model
    # automate_context.store_file_result("./report.pdf")


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    """A function example without inputs.

    If your function does not need any input variables,
     besides what the automation context provides,
     the inputs argument can be omitted.
    """
    pass


# make sure to call the function with the executor
if __name__ == "__main__":
    # NOTE: always pass in the automate function by its reference, do not invoke it!

    # pass in the function reference with the inputs schema to the executor
    execute_automate_function(automate_function, FunctionInputs)

    # if the function has no arguments, the executor can handle it like so
    # execute_automate_function(automate_function_without_inputs)
