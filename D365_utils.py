from config import setup_logger, env
import polars as pl
import requests
from requests.exceptions import HTTPError

log = setup_logger(__name__)
env = env()

class DataverseClient:
    """
    Client for authenticating with and interacting with Microsoft Dataverse.

    Uses OAuth 2.0 client credentials to authenticate against Microsoft Entra ID and provides methods for retrieving and updating Dataverse rows through the Web API.

    Args:
        env: Configuration dictionary containing the Dataverse connection settings. Expected keys are ``CLIENT_ID``, ``CLIENT_SECRET``, ``TENANT_ID``, and ``CRM_URL``.
    """
    def __init__(
            self,
            env : dict = env
            ):

        self.env = env

        self.client_id = env.get("CLIENT_ID", "")
        self.client_secret = env.get("CLIENT_SECRET", "")
        self.tenant_id = env.get("TENANT_ID", "")
        self.crm_url = env.get("CRM_URL", "")

        self.access_token = None
        self.headers = None

    def authenticate(self):
            """
            Fetch and store an access token and auth headers.
            """
    
            log.info("Authenticating with Dataverse")

            token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/token"
    
            token_data = {
                "grant_type" : "client_credentials",
                "client_id" : self.client_id,
                "client_secret" : self.client_secret,
                "resource" : self.crm_url
            }
    
            response = requests.post(token_url, data = token_data)
            response.raise_for_status()
    
            self.access_token = response.json()["access_token"]
    
            self.headers = {
                "Authorization" : f"Bearer {self.access_token}",
                "Accept" : "application/json",
                "OData-MaxVersion" : "4.0",
                "OData-Version" : "4.0",
                "Prefer" : 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
            }
    
            log.info("Authentication successful.")
            
            return self.headers

    def get_row_by_id(
        self,
        table  : str,
        row_id : str,
        select : str | None = None,
        expand : str | None = None,
    ) -> dict:
        """Retrieve a single row from a Dataverse table by its GUID.

        Args:
            table: Logical name of the Dataverse table.
            row_id: GUID of the row to retrieve.
            select: Comma-separated list of columns to return.
            expand: OData ``$expand`` expression for related entities.

        Returns:
            A dictionary containing the requested Dataverse row.

        Raises:
            requests.HTTPError: If the Dataverse request fails.
        """
        if not self.headers:
            log.debug("No active authentication headers. Authenticating.")
            self.authenticate()

        params = []

        if select:
            params.append(f"$select={select}")

        if expand:
            params.append(f"$expand={expand}")

        url = f"{self.crm_url}/api/data/v9.2/{table}({row_id})"

        if params:
            url += "?" + "&".join(params)

        log.debug(
            "Retrieving row '%s' from Dataverse table '%s'.",
            row_id,
            table,
        )

        try:
            response = requests.get(
                url,
                headers = self.headers,
            )
            response.raise_for_status()

        except HTTPError:
            if response.status_code == 401:
                log.warning(
                    "Dataverse authentication expired while retrieving "
                    "row '%s' from '%s'. Refreshing token and retrying.",
                    row_id,
                    table,
                )

                self.authenticate()

                response = requests.get(
                    url,
                    headers=self.headers,
                )
                response.raise_for_status()

            else:
                log.exception(
                    "Failed to retrieve row '%s' from Dataverse table '%s' "
                    "(HTTP %s).",
                    row_id,
                    table,
                    response.status_code,
                )
                raise

        row = response.json()

        log.debug(
            "Successfully retrieved row '%s' from '%s'.",
            row_id,
            table,
        )

        return row

    def list_rows(
        self,
        table: str,
        select: str | None = None,
        filter: str | None = None,
        limit: int | None = None,
        expand: str | None = None
    ) -> list[dict]:
            
            """
        Query rows from a Dataverse table or load dummy data from JSON.

        Builds an OData query string and fetches rows from Dataverse,
        following paging links if necessary.

        Parameters:
            table (str): Name of the Dataverse table to query.
            select (str | None, optional): Comma-separated columns to return. Defaults to None.
            filter (str | None, optional): OData $filter string to filter results. Defaults to None.
            limit (int | None, optional): Maximum number of rows to return. Defaults to None.
            expand (str | None, optional): OData $expand string for related entities. Defaults to None.

        Returns:
            list[dict]: List of rows as dictionaries.
        """


            if not self.headers:
                self.authenticate()

            params = []
            if select:
                params.append(f"$select={select}")
            if filter:
                params.append(f"$filter={filter}")
            if expand:
                params.append(f"$expand={expand}")
            if limit:
                params.append(f"$top={limit}")

            url = f"{self.crm_url}/api/data/v9.2/{table}"
            if params:
                url += "?" + "&".join(params)


            all_rows = []

            while url:
                try:
                    response = requests.get(url, headers=self.headers)
                    response.raise_for_status()
                except HTTPError as e:
                    if response.status_code == 401:
                        log.warning("Token expired. Refreshing...")
                        self.authenticate()
                        response = requests.get(url, headers=self.headers)
                        response.raise_for_status()
                    else:
                        raise e


                data = response.json()
                rows = data.get("value", [])
                all_rows.extend(rows)

                # Follow link to next page if present
                url = data.get("@odata.nextLink")

            return all_rows
        
    def update_row(
        self,
        table : str,
        row_id : str,
        payload : dict[str, object],
    ) -> bool:
        """Update a Dataverse row.

        Args:
            table: Logical name of the Dataverse table.
            row_id: GUID of the row to update.
            payload: Dictionary containing the columns and values to update.

        Returns:
            ``True`` if the row was updated successfully.

        Raises:
            requests.HTTPError: If the update fails after retrying an expired
                authentication token.
        """
        if not self.headers:
            log.debug("No active authentication headers. Authenticating.")
            self.authenticate()

        url = f"{self.crm_url}/api/data/v9.2/{table}({row_id})"

        log.debug(
            "Updating row '%s' on Dataverse table '%s'.",
            row_id,
            table,
        )

        response = requests.patch(
            url,
            headers = self.headers,
            json = payload,
        )

        if response.status_code == 401:
            log.warning(
                "Dataverse authentication expired while updating "
                "row '%s' on table '%s'. Refreshing token and retrying.",
                row_id,
                table,
            )

            self.authenticate()

            response = requests.patch(
                url,
                headers = self.headers,
                json = payload,
            )

        try:
            response.raise_for_status()
        except HTTPError:
            log.exception(
                "Failed to update row '%s' on Dataverse table '%s' "
                "(HTTP %s).",
                row_id,
                table,
                response.status_code,
            )
            raise

        log.debug(
            "Successfully updated row '%s' on Dataverse table '%s'.",
            row_id,
            table,
        )

        return True

    def update_multiple(
        self,
        set_name: str,
        logical_name: str,
        df: pl.DataFrame,
        guid_col: str,
        payload: dict[str, object] | None = None,
        payload_map: dict[str, dict[str, object]] | None = None,
        payload_map_key: str | None = None,
        batch_size: int = 500,
    ) -> None:
        """Update multiple Dataverse rows using the UpdateMultiple operation.

        Rows are submitted to Dataverse in batches. Each row is constructed using
        the supplied entity logical name and GUID column.

        A fixed ``payload`` can be supplied when the same Dataverse fields and
        values should be applied to every row. Alternatively, ``payload_map`` can
        be supplied when the payload varies between rows. Exactly one of
        ``payload`` or ``payload_map`` must be provided.

        If Dataverse returns HTTP 401, the authentication token is refreshed and
        the failed request is retried once.

        Args:
            set_name: Dataverse entity set name used in the Web API endpoint.
                For example, ``contacts`` or ``accounts``.
            logical_name: Dataverse entity logical name. For example,
                ``contact`` or ``account``.
            df: Polars DataFrame containing the rows to update.
            guid_col: Name of the DataFrame column containing the GUID for each
                Dataverse row.
            payload: Fixed mapping of Dataverse field names to values that will be
                applied to every row.
            payload_map: Mapping between values in ``payload_map_key`` and the
                Dataverse fields and values to update. Used when the payload varies
                between rows.
            payload_map_key: Name of the DataFrame column whose value is used to
                select an entry from ``payload_map``.
            batch_size: Maximum number of rows submitted in each request.
                Defaults to 500.

        Returns:
            ``None``. A summary of successful and failed rows is logged when
            processing completes.

        Raises:
            requests.HTTPError: If a batch fails after retrying an expired
                authentication token.
            requests.RequestException: If a network-level error occurs while
                communicating with Dataverse.
            KeyError: If ``guid_col`` is missing from the DataFrame, or if
                ``payload_map_key`` or a corresponding value is missing when using
                ``payload_map``.
            ValueError: If ``batch_size`` is less than 1, or if neither or both
                payload mechanisms are supplied.
        """
        if batch_size < 1:
            raise ValueError(
                "batch_size must be greater than 0. Use the update_row method instead."
            )

        if (payload is None) == (payload_map is None):
            raise ValueError(
                "Provide exactly one of 'payload' or 'payload_map'."
            )

        if payload_map is not None and payload_map_key is None:
            raise ValueError(
                "'payload_map_key' is required when using 'payload_map'."
            )

        if not self.headers:
            log.debug("No active authentication headers. Authenticating.")
            self.authenticate()

        url = (
            f"{self.crm_url}/api/data/v9.2/"
            f"{set_name}/Microsoft.Dynamics.CRM.UpdateMultiple"
        )

        entity_name = f"Microsoft.Dynamics.CRM.{logical_name}"

        successful_count = 0
        failed_count = 0

        rows = list(df.iter_rows(named=True))

        for i in range(0, len(rows), batch_size):
            batch = rows[i:i + batch_size]

            targets = []

            for row in batch:
                row_payload = {
                    "@odata.type": entity_name,
                    f"{logical_name}id": row[guid_col],
                }

                if payload is not None:
                    row_payload.update(payload)
                else:
                    row_payload.update(
                        payload_map[row[payload_map_key]]
                    )

                targets.append(row_payload)

            body = {"Targets": targets}

            log.debug(
                "Updating batch %s-%s in Dataverse table '%s'.",
                i,
                i + len(batch),
                set_name,
            )

            try:
                response = requests.post(
                    url,
                    headers=self.headers,
                    json=body,
                )

                if response.status_code == 401:
                    log.warning(
                        "Dataverse authentication expired while processing "
                        "batch %s-%s. Refreshing token and retrying.",
                        i,
                        i + len(batch),
                    )

                    self.authenticate()

                    response = requests.post(
                        url,
                        headers=self.headers,
                        json=body,
                    )

                response.raise_for_status()

            except HTTPError:
                failed_count += len(batch)

                log.exception(
                    "UpdateMultiple failed for batch %s-%s "
                    "(HTTP %s): %s",
                    i,
                    i + len(batch),
                    response.status_code,
                    response.text,
                )

                raise

            except requests.RequestException:
                failed_count += len(batch)

                log.exception(
                    "Network error while processing UpdateMultiple "
                    "batch %s-%s.",
                    i,
                    i + len(batch),
                )

                raise

            successful_count += len(batch)

        log.info(
            "UpdateMultiple complete. %s succeeded. %s failed.",
            successful_count,
            failed_count,
        )

    def add_row(
    self,
    table: str,
    payload: dict[str, object],
    ) -> str:
        
        """Create a new row in a Dataverse table.

        Args:
            table: Entity set name of the Dataverse table.
                For example, ``contacts`` or ``accounts``.
            payload: Dictionary containing the columns and values for
                the new row.

        Returns:
            The GUID of the newly created Dataverse row.

        Raises:
            requests.HTTPError: If the create request fails after retrying
                an expired authentication token.
        """
        if not self.headers:
            log.debug("No active authentication headers. Authenticating.")
            self.authenticate()

        url = f"{self.crm_url}/api/data/v9.2/{table}"

        log.debug(
            "Creating new row in Dataverse table '%s'.",
            table,
        )

        response = requests.post(
            url,
            headers=self.headers,
            json=payload,
        )

        if response.status_code == 401:
            log.warning(
                "Dataverse authentication expired while creating a row "
                "in table '%s'. Refreshing token and retrying.",
                table,
            )

            self.authenticate()

            response = requests.post(
                url,
                headers=self.headers,
                json=payload,
            )

        try:
            response.raise_for_status()
        except HTTPError:
            log.exception(
                "Failed to create row in Dataverse table '%s' "
                "(HTTP %s): %s",
                table,
                response.status_code,
                response.text,
            )
            raise

            
        entity_id = response.headers.get("OData-EntityId")

        if not entity_id:
            raise RuntimeError(
                "Dataverse created the row but did not return an OData-EntityId."
            )

        row_id = entity_id.rstrip(")").split("(")[-1]

        log.debug(
            "Successfully created row '%s' in Dataverse table '%s'.",
            row_id,
            table,
        )

        return row_id
