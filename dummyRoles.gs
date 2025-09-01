function getDummyRoles() {
  return {
    "roles": [
      {
        "id": 10001,
        "name": "Design",
        "created_at": "2025-01-01T09:00:00Z",
        "updated_at": "2025-06-01T09:00:00Z",
        "user_ids": [101, 102, 103, 104]
      },
      {
        "id": 10002,
        "name": "Engineering",
        "created_at": "2025-01-02T09:00:00Z",
        "updated_at": "2025-06-02T09:00:00Z",
        "user_ids": [201, 202, 203, 204, 205]
      },
      {
        "id": 10003,
        "name": "Marketing",
        "created_at": "2025-01-03T09:00:00Z",
        "updated_at": "2025-06-03T09:00:00Z",
        "user_ids": [301, 302, 303]
      }
    ],
    "per_page": 2000,
    "total_pages": 1,
    "total_entries": 3,
    "next_page": null,
    "previous_page": null,
    "page": 1,
    "links": {
      "first": "https://dummy.harvestapp.com/v2/roles?page=1&per_page=2000&ref=first",
      "next": null,
      "previous": null,
      "last": "https://dummy.harvestapp.com/v2/roles?page=1&per_page=2000&ref=last"
    }
  };
}