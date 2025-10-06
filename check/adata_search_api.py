import requests
import logging
from typing import Dict, List, Optional, Any
from datetime import datetime

class AdataAPI:
    '''An API agent for interacting with the adata.kz'''

    def __init__(
            self,
            logger: Optional[logging.Logger] = None,
            DO_LOGGING: bool = True,
        ):
        self.logger = logger if logger else logging.getLogger(__name__)
        self.DO_LOGGING = DO_LOGGING
        self.base_url = "https://pk-api.adata.kz/api/v1"
        self.user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

    def search(self, keyword: str) -> Dict[str, Any]:
        """
        Search for companies by keyword

        Args:
            keyword (str): Search keyword (company name, director name, etc.)

        Returns:
            Dict with search results
        """
        url = f"{self.base_url}/data/search"
        headers = {"User-Agent": self.user_agent}
        params = {
            "most_viewed_companies": 0,
            "keyword": keyword,
        }

        try:
            if self.DO_LOGGING and self.logger:
                self.logger.debug(f"Request to: {url}, with params: {params}")

            response = requests.get(
                headers=headers,
                params=params,
                url=url,
                timeout=10
            )
            response.raise_for_status()
            return response.json()

        except requests.exceptions.RequestException as e:
            if self.DO_LOGGING and self.logger:
                self.logger.error(f"Request failed: {e}")
            return {"status": False, "error": str(e)}

    def get_company_by_biin(self, biin: str) -> Optional[Dict[str, Any]]:
        """
        Get company details by BIIN (БИИН)

        Args:
            biin (str): Company BIIN number

        Returns:
            Company details or None if not found
        """
        # First search by BIIN
        result = self.search(biin)

        if result.get('status') and result['data']['count_all'] > 0:
            # Try to find exact BIIN match
            for company in result['data']['result']:
                if company.get('biin') == biin:
                    return company
        return None

    def extract_company_info(self, company_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Extract and format company information from raw API response

        Args:
            company_data (Dict): Raw company data from API

        Returns:
            Formatted company information
        """
        return {
            'id': company_data.get('id'),
            'biin': company_data.get('biin'),
            'name': company_data.get('name'),
            'address': company_data.get('address'),
            'trustworthy': company_data.get('trustworthy', False),
            'type_id': company_data.get('type_id'),
            'is_inactive': company_data.get('is_inactive', False),
            'registration_date': company_data.get('registration_date'),
            'director_name': company_data.get('director_name'),
            'status': company_data.get('status'),
            'status_code': company_data.get('status_code'),
            'highlight': company_data.get('highlight', []),
            'score': company_data.get('_score')
        }

    def search_companies(self, keyword: str, max_results: int = 10) -> List[Dict[str, Any]]:
        """
        Search for companies and return formatted results

        Args:
            keyword (str): Search keyword
            max_results (int): Maximum number of results to return

        Returns:
            List of formatted company information
        """
        result = self.search(keyword)

        if not result.get('status'):
            if self.DO_LOGGING and self.logger:
                self.logger.warning(f"Search failed for keyword: {keyword}")
            return []

        companies = []
        for company in result['data']['result'][:max_results]:
            companies.append(self.extract_company_info(company))

        return companies

    def get_companies_by_director(self, director_name: str) -> List[Dict[str, Any]]:
        """
        Search for companies by director name

        Args:
            director_name (str): Director's full or partial name

        Returns:
            List of companies associated with the director
        """
        result = self.search(director_name)

        if not result.get('status'):
            return []

        director_companies = []
        for company in result['data']['result']:
            # Check if director name matches or is in highlights
            company_director = company.get('director_name', '').lower()
            search_director = director_name.lower()

            if (search_director in company_director or
                any(search_director in str(h).lower() for h in company.get('highlight', []))):
                director_companies.append(self.extract_company_info(company))

        return director_companies

    def is_company_active(self, biin_or_name: str) -> bool:
        """
        Check if a company is active

        Args:
            biin_or_name (str): Company BIIN or name

        Returns:
            True if company exists and is active, False otherwise
        """
        result = self.search(biin_or_name)

        if result.get('status') and result['data']['count_all'] > 0:
            for company in result['data']['result']:
                if (company.get('biin') == biin_or_name or
                    company.get('name', '').lower() == biin_or_name.lower()):
                    return not company.get('is_inactive', False)
        return False

# Usage examples
if __name__ == "__main__":
    # Setup logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)

    # Initialize API client
    api = AdataAPI(logger=logger, DO_LOGGING=True)

    # Example 1: Search by company name or director
    print("=== Поиск компаний ===")
    companies = api.search_companies("Серединский Эрнесто")
    for company in companies:
        print(f"Компания: {company['name']}")
        print(f"Директор: {company['director_name']}")
        print(f"Статус: {company['status']}")
        print(f"БИИН: {company['biin']}")
        print("-" * 50)

    # Example 2: Get company by BIIN
    print("\n=== Поиск по БИИН ===")
    company = api.get_company_by_biin("170740005168")
    if company:
        print(f"Найдена компания: {company['name']}")

    # Example 3: Check company status
    print("\n=== Проверка статуса компании ===")
    is_active = api.is_company_active("170740005168")
    print(f"Компания активна: {is_active}")

    # Example 4: Search by director name
    print("\n=== Поиск по директору ===")
    director_companies = api.get_companies_by_director("ГУЛИМБЕТОВ ДОСБОЛ")
    for comp in director_companies:
        print(f"{comp['name']} - {comp['director_name']}")
