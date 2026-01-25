from PyQt6.QtGui import QUndoCommand


class AddObjectCommand(QUndoCommand):
    def __init__(self, model, obj, description: str = "Add Object") -> None:
        super().__init__(description)
        self.model = model
        self.obj = obj

    def redo(self) -> None:
        self.model.add_object(self.obj)

    def undo(self) -> None:
        self.model.remove_object(self.obj.id)


class RemoveObjectCommand(QUndoCommand):
    def __init__(self, model, obj, description: str = "Remove Object") -> None:
        super().__init__(description)
        self.model = model
        self.obj = obj

    def redo(self) -> None:
        self.model.remove_object(self.obj.id)

    def undo(self) -> None:
        self.model.add_object(self.obj)


class UpdateObjectCommand(QUndoCommand):
    def __init__(self, model, old_obj, new_obj, description: str) -> None:
        super().__init__(description)
        self.model = model
        self.old_obj = old_obj
        self.new_obj = new_obj

    def redo(self) -> None:
        self.model.update_object(self.old_obj.id, self.new_obj)

    def undo(self) -> None:
        self.model.update_object(self.old_obj.id, self.old_obj)


class UpdateClassificationCommand(QUndoCommand):
    def __init__(
        self,
        model,
        old_text: str,
        old_size: int,
        new_text: str,
        new_size: int,
        description: str = "Update Classification",
    ) -> None:
        super().__init__(description)
        self.model = model
        self.old_text = old_text
        self.old_size = old_size
        self.new_text = new_text
        self.new_size = new_size

    def redo(self) -> None:
        self.model.set_classification(self.new_text, self.new_size)

    def undo(self) -> None:
        self.model.set_classification(self.old_text, self.old_size)


class AddTopicCommand(QUndoCommand):
    def __init__(self, model, topic, index: int | None = None, description: str = "Add Topic") -> None:
        super().__init__(description)
        self.model = model
        self.topic = topic
        self.index = index

    def redo(self) -> None:
        self.model.insert_topic(self.topic, self.index)

    def undo(self) -> None:
        self.model.remove_topic(self.topic.id)


class UpdateTopicCommand(QUndoCommand):
    def __init__(self, model, old_topic, new_topic, description: str = "Update Topic") -> None:
        super().__init__(description)
        self.model = model
        self.old_topic = old_topic
        self.new_topic = new_topic

    def redo(self) -> None:
        self.model.update_topic(self.old_topic.id, self.new_topic)

    def undo(self) -> None:
        self.model.update_topic(self.old_topic.id, self.old_topic)


class UpdateDeliverableCommand(QUndoCommand):
    def __init__(
        self, model, old_deliverable, new_deliverable, description: str = "Update Deliverable"
    ) -> None:
        super().__init__(description)
        self.model = model
        self.old_deliverable = old_deliverable
        self.new_deliverable = new_deliverable

    def redo(self) -> None:
        self.model.update_deliverable(self.old_deliverable.id, self.new_deliverable)

    def undo(self) -> None:
        self.model.update_deliverable(self.old_deliverable.id, self.old_deliverable)


class ToggleTopicCollapseCommand(QUndoCommand):
    def __init__(self, model, topic_id: str, was_collapsed: bool) -> None:
        super().__init__("Toggle Topic Collapse")
        self.model = model
        self.topic_id = topic_id
        self.was_collapsed = was_collapsed

    def redo(self) -> None:
        self.model.toggle_topic_collapsed(self.topic_id)

    def undo(self) -> None:
        topic = self.model.get_topic(self.topic_id)
        if topic is None:
            return
        if topic.collapsed != self.was_collapsed:
            self.model.toggle_topic_collapsed(self.topic_id)


class AddDeliverableCommand(QUndoCommand):
    def __init__(self, model, topic_id: str, deliverable, index: int | None = None) -> None:
        super().__init__("Add Deliverable")
        self.model = model
        self.topic_id = topic_id
        self.deliverable = deliverable
        self.index = index

    def redo(self) -> None:
        self.model.insert_deliverable(self.topic_id, self.deliverable, self.index)

    def undo(self) -> None:
        self.model.remove_deliverable(self.deliverable.id)


class MoveDeliverableCommand(QUndoCommand):
    def __init__(self, model, deliverable_id: str, old_index: int, new_index: int) -> None:
        super().__init__("Move Deliverable")
        self.model = model
        self.deliverable_id = deliverable_id
        self.old_index = old_index
        self.new_index = new_index

    def redo(self) -> None:
        self.model.move_deliverable(self.deliverable_id, self.new_index)

    def undo(self) -> None:
        self.model.move_deliverable(self.deliverable_id, self.old_index)


class MoveDeliverableAcrossTopicsCommand(QUndoCommand):
    def __init__(
        self,
        model,
        deliverable_id: str,
        source_topic_id: str,
        source_index: int,
        target_topic_id: str,
        target_index: int,
    ) -> None:
        super().__init__("Move Deliverable")
        self.model = model
        self.deliverable_id = deliverable_id
        self.source_topic_id = source_topic_id
        self.source_index = source_index
        self.target_topic_id = target_topic_id
        self.target_index = target_index

    def redo(self) -> None:
        self.model.move_deliverable_to_topic(
            self.deliverable_id, self.target_topic_id, self.target_index
        )

    def undo(self) -> None:
        self.model.move_deliverable_to_topic(
            self.deliverable_id, self.source_topic_id, self.source_index
        )


class RemoveDeliverableCommand(QUndoCommand):
    def __init__(self, model, topic_id: str, deliverable, index: int, removed_objects) -> None:
        super().__init__("Remove Deliverable")
        self.model = model
        self.topic_id = topic_id
        self.deliverable = deliverable
        self.index = index
        self.removed_objects = removed_objects

    def redo(self) -> None:
        for obj in self.removed_objects:
            self.model.remove_object(obj.id)
        self.model.remove_deliverable(self.deliverable.id)

    def undo(self) -> None:
        self.model.insert_deliverable(self.topic_id, self.deliverable, self.index)
        for obj in self.removed_objects:
            self.model.add_object(obj)


class RemoveTopicCommand(QUndoCommand):
    def __init__(self, model, topic, index: int, removed_objects) -> None:
        super().__init__("Remove Topic")
        self.model = model
        self.topic = topic
        self.index = index
        self.removed_objects = removed_objects

    def redo(self) -> None:
        for obj in self.removed_objects:
            self.model.remove_object(obj.id)
        self.model.remove_topic(self.topic.id)

    def undo(self) -> None:
        self.model.insert_topic(self.topic, self.index)
        for obj in self.removed_objects:
            self.model.add_object(obj)
