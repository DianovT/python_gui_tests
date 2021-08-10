def test_delete_group(app):
    old_list = app.groups.get_group_list()
    app.groups.delete_group("my group")
    new_list = app.groups.get_group_list()
    old_list.remove()
    assert sorted(old_list) == sorted(new_list)
