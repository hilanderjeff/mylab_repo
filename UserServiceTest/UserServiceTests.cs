using Moq;
using Xunit;
using Microsoft.Extensions.Logging;
using UserServiceMain;

namespace UserServiceTest
{
    public class UserServiceTests
    {
        private readonly Mock<IUserRepository> _mockRepository;
        private readonly Mock<ILogger<UserService>> _mockLogger;
        private readonly IUserService _userService;

        public UserServiceTests()
        {
            _mockRepository = new Mock<IUserRepository>();
            _mockLogger = new Mock<ILogger<UserService>>();
            _userService = new UserService(_mockRepository.Object, _mockLogger.Object);
        }
        [Fact]
        public async Task GetUserById_ExistingUser_ReturnsUser()
        {
            // Arrange
            var expectedUser = new User { Id = 1, Name = "John Doe", Email = "john@example.com" };
            _mockRepository.Setup(repo => repo.GetByIdAsync(1))
                .ReturnsAsync(expectedUser);

            // Act
            var result = await _userService.GetUserByIdAsync(1);

            // Assert
            //Assert.NotNull(result);
            Assert.Equal(expectedUser.Id, result.Id);
            //Assert.Equal(expectedUser.Name, result.Name);
            //_mockRepository.Verify(repo => repo.GetByIdAsync(2), Times.Once);
        }

        [Fact]
        public async Task GetUserById_NonExistingUser_ThrowsNotFoundException()
        {
            // Arrange
            _mockRepository.Setup(repo => repo.GetByIdAsync(It.IsAny<int>()))
                .ReturnsAsync((User)null);

            // Act & Assert
            await Assert.ThrowsAsync<NotFoundException>(() =>
                _userService.GetUserByIdAsync(1));

            _mockRepository.Verify(repo => repo.GetByIdAsync(1), Times.Exactly(1));
        }

        [Fact]
        public async Task GetActiveUsers_ReturnsOnlyUsersWithEmail()
        {
            // Arrange
            var users = new List<User>
            {
                new User { Id = 1, Name = "John", Email = "john@example.com" },
                new User { Id = 2, Name = "Jane", Email = "" },
                new User { Id = 3, Name = "Bob", Email = "bob@example.com" }
            };

            _mockRepository.Setup(repo => repo.GetAllAsync())
                .ReturnsAsync(users);

            // Act
            var result = await _userService.GetActiveUsersAsync();

            // Assert
            Assert.Equal(2, result.Count());
            Assert.All(result, user => Assert.False(string.IsNullOrEmpty(user.Email)));
            _mockRepository.Verify(repo => repo.GetAllAsync(), Times.Once);
        }

    }
}