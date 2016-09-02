package com.dvoinov.parsexlscards;

import java.util.Date;

/**
 * Created by dvoinov on 31.08.2016.
 */
public class Card {
    private int cardId;
    private String cardHolder;
    private boolean accEnter1;
    private boolean accExit1;
    private boolean accEnter2;
    private boolean accExit2;
    private boolean accServ;
    private boolean accBuhg;
    private boolean accKlad;
    private boolean accDir;
    private Date lastAccess;
    private boolean blocked;

    public boolean isBlocked() {
        return blocked;
    }

    public void setBlocked(boolean blocked) {
        this.blocked = blocked;
    }

    public int getCardId() {
        return cardId;
    }

    public Date getLastAccess() {
        return lastAccess;
    }

    public void setLastAccess(Date lastAccess) {
        this.lastAccess = lastAccess;
    }

    public void setCardId(int cardId) {
        this.cardId = cardId;
    }

    public String getCardHolder() {
        return cardHolder;
    }

    public void setCardHolder(String cardHolder) {
        this.cardHolder = cardHolder;
    }

    public boolean isAccEnter1() {
        return accEnter1;
    }

    public void setAccEnter1(boolean accEnter1) {
        this.accEnter1 = accEnter1;
    }

    public boolean isAccExit1() {
        return accExit1;
    }

    public void setAccExit1(boolean accExit1) {
        this.accExit1 = accExit1;
    }

    public boolean isAccEnter2() {
        return accEnter2;
    }

    public void setAccEnter2(boolean accEnter2) {
        this.accEnter2 = accEnter2;
    }

    public boolean isAccExit2() {
        return accExit2;
    }

    public void setAccExit2(boolean accExit2) {
        this.accExit2 = accExit2;
    }

    public boolean isAccServ() {
        return accServ;
    }

    public void setAccServ(boolean accServ) {
        this.accServ = accServ;
    }

    public boolean isAccBuhg() {
        return accBuhg;
    }

    public void setAccBuhg(boolean accBuhg) {
        this.accBuhg = accBuhg;
    }

    public boolean isAccKlad() {
        return accKlad;
    }

    public void setAccKlad(boolean accKlad) {
        this.accKlad = accKlad;
    }

    public boolean isAccDir() {
        return accDir;
    }

    public void setAccDir(boolean accDir) {
        this.accDir = accDir;
    }
}
