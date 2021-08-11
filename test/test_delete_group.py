import random


def test_delete_group(app):
    if len(app.groups.get_group_list()) == 1:
        app.groups.add_new_group("my group")
    old_list = app.groups.get_group_list()
    group = random.choice(old_list)
    index = old_list.index(group)
    app.groups.delete_group(index)
    new_list = app.groups.get_group_list()
    old_list.remove(group)
    assert sorted(old_list) == sorted(new_list)
