from infrastructure.db.queries import insert_template_json, get_template_json_by_id

def save_template(template_id, template_json):
    insert_template_json(template_id, template_json)
    # add save logger

