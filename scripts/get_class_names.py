def list_ci_class_names():
    ci_url = f"{instance}/api/now/table/cmdb_ci"
    params = {
        "sysparm_fields": "sys_class_name",
        "sysparm_limit": "10000",
        "sysparm_display_value": "false"
    }

    res = requests.get(ci_url, auth=HTTPBasicAuth(username, password), params=params)
    if res.status_code == 200:
        records = res.json().get("result", [])
        classes = set(r.get("sys_class_name") for r in records if r.get("sys_class_name"))
        print("🔍 Unique CI Classes:")
        for c in sorted(classes):
            print(" -", c)
    else:
        print(f"❌ Failed to fetch: {res.status_code}")
        print(res.text)
