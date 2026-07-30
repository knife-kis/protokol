package ru.citlab24.protokol.tabs.models;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Building implements Serializable {
    private static final long serialVersionUID = 1L;
    private int plannedFloorsCount;
    private int id;
    private String name;
    private int revision = 1;
    private int sourceProjectId;
    private String createdAt;
    private String updatedAt;
    private String createdBy;
    private String updatedBy;
    private String lockOwner;
    private TitlePageData titlePageData = new TitlePageData();
    private final List<Section> sections = new ArrayList<>(); // <-- секции
    private final List<Floor> floors = new ArrayList<>();

    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public int getRevision() { return revision; }
    public void setRevision(int revision) { this.revision = revision; }

    public int getSourceProjectId() { return sourceProjectId; }
    public void setSourceProjectId(int sourceProjectId) { this.sourceProjectId = sourceProjectId; }

    public String getCreatedAt() { return createdAt; }
    public void setCreatedAt(String createdAt) { this.createdAt = createdAt; }

    public String getUpdatedAt() { return updatedAt; }
    public void setUpdatedAt(String updatedAt) { this.updatedAt = updatedAt; }

    public String getCreatedBy() { return createdBy; }
    public void setCreatedBy(String createdBy) { this.createdBy = createdBy; }

    public String getUpdatedBy() { return updatedBy; }
    public void setUpdatedBy(String updatedBy) { this.updatedBy = updatedBy; }

    public String getLockOwner() { return lockOwner; }
    public void setLockOwner(String lockOwner) { this.lockOwner = lockOwner; }

    public List<Floor> getFloors() { return floors; }
    public void addFloor(Floor floor) { floors.add(floor); }


    // секции
    public List<Section> getSections() { return sections; }
    public void setSections(List<Section> newSections) {
        sections.clear();
        if (newSections != null) sections.addAll(newSections);
    }
    public void addSection(Section s) { sections.add(s); }

    public TitlePageData getTitlePageData() {
        if (titlePageData == null) {
            titlePageData = new TitlePageData();
        }
        return titlePageData;
    }

    public void setTitlePageData(TitlePageData titlePageData) {
        this.titlePageData = (titlePageData == null) ? new TitlePageData() : titlePageData;
    }
}
