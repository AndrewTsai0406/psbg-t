style_transfer_prompt_question_mark = """
You are helping me rewrite the template using a [Template] as a guide and multiple [Contexts] as references.

The [Template] may contain one or more '?' that need to be replaced with numerical values from the [Context].
As an editor, you need to maintain the format of the [Template] and only replace the '?' with numerical values from the [Context]. 

Strictly follow the [Template] format. If there is no '?' to be changed in the [Template], don't change anything.'
Output the texts directly after [Revised Template]. [List_Items] is a list of items that indicates a broader context for the [Template]. You can use this information to help you replace the values in the [Template].

Note that the [Context] does not necessarily contain all the values needed to replace the '?' in the [Template].
Be very careful to only replace the '?' with the correct values if it can be derived from the [Context].

You must respond using JSON format, according to the following schema:
{arg_schema}

[Context]:{context}

[List_Items]:{List_Items}

[Template]:{template}

"""

style_transfer_prompt_question_mark_v2 = """
I will provide you with multiple [Contexts] containing hierarchical headings, along with a [Template] and the corresponding hierarchical headings of the [Template] as [Hierarchical_Heading]. Please rewrite the [Template] based on the content of the [Contexts].

Please follow these rules when revising:

* The [Template] may contain one or more '?' that need to be replaced with numerical values from the [Context]. As an editor, you need to maintain the format of the [Template] and only replace the '?' with numerical values from the [Context].

* Strictly follow the format of the [Template]. If the information provided in the [Context] cannot be used to revise the [Template], return the original [Template] as the final result.

* You must respond using JSON format, according to the following schema:\n{arg_schema}

* Place the context numbers you referenced for the revision in 'context_paragraph_ids' and the revised result in 'revised_template', and avoid giving any explanations.
------------------------------------------------------------------------------------------------------------------------------------------
[Context]:{context}

[Hierarchical_Heading]:{List_Items}

[Template]:{template}
"""


table_style_transfer_prompt = """
Task: Replace value with new value based on [Context].

Using the provided [Context], replace the values in the given row while keeping the same format as the original [Template]. 
Be sure to use the new values appropriately for each column. [List_Items] is a list of items that indicates a broader context for the [Template]. You can use this information to help you replace the values in the [Template].
Note that the number of rows in each Column in the [Template] and [Revised Template] are not necessarily the same.
Output only the dictionary that can be turned to a dataframe directly after [Revised Template].

Note that the [Context] does not necessarily contain all the values needed to replace the '?' in the [Template].
Be very careful to only replace the '?' with the correct values if it can be derived from the [Context].

For example:

    [Template]:{{Column A: {{0: value_a1, 1: value_a2}}, Column B: {{0: value_b1, 1: value_b2}}, Column C: {{0: value_c1, 1: value_c2}}}}
    [Revised Template]:{{Column A: {{0: new_value_a1, 1: new_value_a2}}, Column B: {{0: new_value_b1, 1: new_value_b2}}, Column C: {{0: new_value_c1, 1: new_value_c2}}}}

[Context]:
{context}

[List_Items]:
{List_Items}

[Template]:
For reference only, please replace the ? in the following table with the correct values based on the [Context]:
{template}

"""


table_style_transfer_prompt_v2 = """
Task: Replace values with new values based on [Context].

<Instructions>
* Context Utilization
    Use the provided [Context] to replace the values in the given row.
    Maintain the same format as the original [Template].
* Value Replacement
    Ensure that new values are appropriately applied to each column.
    Replace each '?' with a value derived from the [Context] if available; otherwise, leave the '?' unchanged.
    The [Context] might not contain all necessary values to replace every ? in the [Template].
    Exercise caution to ensure that only replaceable ? are updated based on the [Context].
* List_Items Consideration
    [List_Items] serves as a broader context for the [Template].
    Utilize this list to assist in accurately replacing values within the [Template].
* Template Structure
    The number of rows in each column of the [Template] and [Revised Template] may differ.
* Output Requirements
    Output only the dictionary formatted for direct conversion into a dataframe after [Revised Template].


For example:

   [Template]:{{Column A: {{0: ?, 1: ?}}, Column B: {{0: ?, 1: ?}}, Column C: {{0: ?, 1: ?}}}}
   [Revised Template]:{{Column A: {{0: new_value_a1, 1: new_value_a2}}, Column B: {{0: new_value_b1, 1: new_value_b2}}, Column C: {{0: new_value_c1, 1: new_value_c2}}}}

------------------------------------------------------------------------------------------------

Below is an example of the target table, this is for REFERENCE ONLY. You should still replace the “?” in the [Template] with the correct values based on the [Context]. Don’t copy from the example table.
{content}

[Template]:
{content_masked}

[List_Items]:
{List_Items}

[Context]:
{context}
"""
