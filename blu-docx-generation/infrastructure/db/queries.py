from infrastructure.db.query_executor import fetch_one, execute

def get_template_json_by_id(template_id: str):
    query = """
    SELECT template_id, template_json
    FROM docx_json
    WHERE template_id = %s
    """

    return fetch_one(query, (template_id))

def insert_template_json(template_id ,template_json):
    query = """
    INSERT INTO docx_json (template_id, template_json)
    values (%s, %s)
    ON DUPLICATE KEY UPDATE
        template_json = VALUES(template_json)

    """
    execute(
        query,
        (
            template_id, template_json
        )
    )