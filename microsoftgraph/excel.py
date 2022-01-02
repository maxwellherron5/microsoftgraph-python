import base64
import mimetypes

from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Excel(object):
    def __init__(self, client) -> None:
        """Use the Excel REST API

        https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client
    
    @token_required
    def create_session(self, bar):
        """Create a workbook session
        """
        ...
    
    @token_required
    def calculate_workbook(self, bar):
        """Calculates a given workbook session

        https://docs.microsoft.com/en-us/graph/api/workbookapplication-calculate?view=graph-rest-1.0&tabs=http

        Args:
            bar (str): Unique identifier for the session.

        Returns:
            Response: Microsoft Graph Response.
        """
        ...
    
    @token_required
    def list_tables(self, bar) -> Response:
        """Calculates a given workbook session

        https://docs.microsoft.com/en-us/graph/api/workbookapplication-calculate?view=graph-rest-1.0&tabs=http

        Args:
            bar (str): Unique identifier for the session.

        Returns:
            Response: Microsoft Graph Response.
        """
        ...
    
    @token_required
    def list_worksheets(self, bar) -> Response:
        """Calculates a given workbook session

        https://docs.microsoft.com/en-us/graph/api/workbookapplication-calculate?view=graph-rest-1.0&tabs=http

        Args:
            bar (str): Unique identifier for the session.

        Returns:
            Response: Microsoft Graph Response.
        """
        ...

    @token_required
    def list_names(self, bar) -> Response:
        """Calculates a given workbook session

        https://docs.microsoft.com/en-us/graph/api/workbookapplication-calculate?view=graph-rest-1.0&tabs=http

        Args:
            bar (str): Unique identifier for the session.

        Returns:
            Response: Microsoft Graph Response.
        """
        ...

    @token_required
    def get_operation_result(self, bar) -> Response:
        """Calculates a given workbook session

        https://docs.microsoft.com/en-us/graph/api/workbookapplication-calculate?view=graph-rest-1.0&tabs=http

        Args:
            bar (str): Unique identifier for the session.

        Returns:
            Response: Microsoft Graph Response.
        """
        ...
    
    @token_required
    def calculate_workbook(self, id: str, calculation_type: str = "FullRebuild") -> Response:
        """Calculates a given workbook session

        https://docs.microsoft.com/en-us/graph/api/workbookapplication-calculate?view=graph-rest-1.0&tabs=http

        Args:
            bar (str): Unique identifier for the session.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = f"me/drive/items/{id}/workbook/application/calculate"
        calculate_msg = {
            "CalculationType": calculation_type,
        }

        return self._client._post(self._client.base_url + url, json=calculate_msg)
