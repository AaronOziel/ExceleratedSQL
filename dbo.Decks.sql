CREATE TABLE [dbo].[Decks] (
    [Deck Name]    VARCHAR (50) NOT NULL,
    [Date Created] DATE     NULL,
    [Rank]         SMALLINT     NULL,
    [Player Name]  VARCHAR (50) NOT NULL,
    [ID]           INT          NOT NULL,
    [Pro]          BIT          NOT NULL,
    PRIMARY KEY CLUSTERED ([ID] ASC)
);

